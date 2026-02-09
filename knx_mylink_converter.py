#!/usr/bin/env python3
"""Convert ETS group address export for MaiLian device #9 (KNX DALI tunable white dimmer).

Outputs:
1) Renamed XML file
2) XLSX mapping table for import/checking
"""

from __future__ import annotations

import argparse
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple
from zipfile import ZIP_DEFLATED, ZipFile
import xml.sax.saxutils as saxutils

NS_URI = "http://knx.org/xml/ga-export/01"
NS = {"k": NS_URI}
ET.register_namespace("", NS_URI)

# Device mapping recovered from the provided MaiLian mapping table for device 9.
FUNCTIONS_BY_MIDDLE: Dict[int, Tuple[str, int]] = {
    1: ("开关写", 3),
    2: ("亮度写", 1),
    3: ("色温写", 5),
    4: ("开关读", 2),
    5: ("亮度读", 0),
    6: ("色温读", 4),
}
DEVICE_NO = 9
BLOCK_SIZE = 80
LIGHTS_PER_BLOCK = 64
GROUPS_PER_BLOCK = 16
BLOCKS_PER_MAIN = 2


@dataclass
class Row:
    name: str
    address: str
    main: int
    middle: int
    sub: int
    module_global: int
    module_in_main: int
    object_type: str
    object_no: int
    device_no: int
    func_no: int
    func_name: str


def parse_address(address: str) -> Tuple[int, int, int]:
    parts = address.split("/")
    if len(parts) != 3:
        raise ValueError(f"Invalid group address: {address}")
    return int(parts[0]), int(parts[1]), int(parts[2])


def classify_object(sub: int) -> Tuple[int, str, int]:
    if sub < 1:
        raise ValueError("Sub group must be >= 1")
    block_idx = (sub - 1) // BLOCK_SIZE + 1
    local = (sub - 1) % BLOCK_SIZE + 1
    if local <= LIGHTS_PER_BLOCK:
        return block_idx, "灯", local
    return block_idx, "组", local - LIGHTS_PER_BLOCK


def build_new_name(object_type: str, object_no: int, func_no: int) -> str:
    return f"{object_type}{object_no} {DEVICE_NO} {func_no}"


def should_convert(main: int, middle: int, sub: int) -> bool:
    if main < 1 or main > 4:
        return False
    if middle not in FUNCTIONS_BY_MIDDLE:
        return False
    # Current project rule: 2 modules per main group -> sub 1..160
    return 1 <= sub <= BLOCK_SIZE * BLOCKS_PER_MAIN


def convert(xml_input: Path, xml_output: Path, xlsx_output: Path) -> Tuple[int, int]:
    tree = ET.parse(xml_input)
    root = tree.getroot()

    rows: List[Row] = []
    converted = 0

    for ga in root.findall(".//k:GroupAddress", NS):
        address = ga.get("Address", "")
        if not address:
            continue
        try:
            main, middle, sub = parse_address(address)
        except ValueError:
            continue

        if not should_convert(main, middle, sub):
            continue

        func_name, func_no = FUNCTIONS_BY_MIDDLE[middle]
        module_in_main, object_type, object_no = classify_object(sub)
        module_global = (main - 1) * BLOCKS_PER_MAIN + module_in_main

        new_name = build_new_name(object_type, object_no, func_no)
        ga.set("Name", new_name)
        converted += 1

        rows.append(
            Row(
                name=new_name,
                address=address,
                main=main,
                middle=middle,
                sub=sub,
                module_global=module_global,
                module_in_main=module_in_main,
                object_type=object_type,
                object_no=object_no,
                device_no=DEVICE_NO,
                func_no=func_no,
                func_name=func_name,
            )
        )

    tree.write(xml_output, encoding="utf-8", xml_declaration=True)
    write_xlsx(xlsx_output, rows)
    return converted, len(rows)


def excel_col(col: int) -> str:
    out = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        out = chr(ord("A") + rem) + out
    return out


def cell_ref(row_idx: int, col_idx: int) -> str:
    return f"{excel_col(col_idx)}{row_idx}"


def inline_cell(row_idx: int, col_idx: int, value: str) -> str:
    ref = cell_ref(row_idx, col_idx)
    return (
        f'<c r="{ref}" t="inlineStr">'
        f"<is><t>{saxutils.escape(value)}</t></is>"
        f"</c>"
    )


def write_xlsx(path: Path, rows: List[Row]) -> None:
    headers = [
        "Name",
        "Address",
        "Main",
        "Middle",
        "Sub",
        "ModuleGlobal",
        "ModuleInMain",
        "ObjectType",
        "ObjectNo",
        "DeviceNo",
        "FunctionNo",
        "FunctionName",
    ]

    all_rows: List[List[str]] = [headers]
    for r in rows:
        all_rows.append(
            [
                r.name,
                r.address,
                str(r.main),
                str(r.middle),
                str(r.sub),
                str(r.module_global),
                str(r.module_in_main),
                r.object_type,
                str(r.object_no),
                str(r.device_no),
                str(r.func_no),
                r.func_name,
            ]
        )

    sheet_rows: List[str] = []
    for r_idx, row in enumerate(all_rows, start=1):
        cells = "".join(inline_cell(r_idx, c_idx, val) for c_idx, val in enumerate(row, start=1))
        sheet_rows.append(f'<row r="{r_idx}">{cells}</row>')

    dimension = f"A1:{excel_col(len(headers))}{len(all_rows)}"

    worksheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"<dimension ref=\"{dimension}\"/>"
        "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>"
        "<sheetFormatPr defaultRowHeight=\"15\"/>"
        f"<sheetData>{''.join(sheet_rows)}</sheetData>"
        "</worksheet>"
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheets><sheet name=\"Converted\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
        "</workbook>"
    )

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        "</Types>"
    )

    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
        'Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
        'Target="docProps/app.xml"/>'
        "</Relationships>"
    )

    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    )

    core_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:creator>Codex Converter</dc:creator>"
        "<cp:lastModifiedBy>Codex Converter</cp:lastModifiedBy>"
        "</cp:coreProperties>"
    )

    app_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>Codex</Application>"
        "</Properties>"
    )

    with ZipFile(path, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("docProps/core.xml", core_xml)
        zf.writestr("docProps/app.xml", app_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        zf.writestr("xl/worksheets/sheet1.xml", worksheet_xml)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert ETS group addresses for MaiLian device #9 and export XML + XLSX"
    )
    parser.add_argument("--xml-input", required=True, help="Input ETS XML path")
    parser.add_argument("--xml-output", required=True, help="Output converted XML path")
    parser.add_argument("--xlsx-output", required=True, help="Output XLSX path")
    args = parser.parse_args()

    xml_input = Path(args.xml_input)
    xml_output = Path(args.xml_output)
    xlsx_output = Path(args.xlsx_output)

    converted, rows = convert(xml_input, xml_output, xlsx_output)
    print(f"Converted group addresses: {converted}")
    print(f"XLSX rows (without header): {rows}")
    print(f"XML output: {xml_output}")
    print(f"XLSX output: {xlsx_output}")


if __name__ == "__main__":
    main()
