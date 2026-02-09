#!/usr/bin/env python3
"""Local web tool for KNX ETS XML rename + XML/XLSX export.

Features:
- Upload ETS XML in browser
- Auto-generate names for MaiLian device #9
- Edit any name in web table
- Batch rename by CSV template
- Export modified XML and XLSX
"""

from __future__ import annotations

import argparse
import io
import json
import math
import re
import threading
import uuid
import webbrowser
from dataclasses import dataclass, asdict
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Dict, List, Tuple
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET
import xml.sax.saxutils as saxutils

NS_URI = "http://knx.org/xml/ga-export/01"
NS = {"k": NS_URI}
ET.register_namespace("", NS_URI)

PROJECT_NAME_ZH = "KNX 群组地址转换工作台"
PROJECT_NAME_EN = "KNX Group Address Studio"

DEVICE_NO = 9
FUNCTIONS_BY_MIDDLE: Dict[int, Tuple[str, int]] = {
    1: ("开关写", 3),
    2: ("亮度写", 1),
    3: ("色温写", 5),
    4: ("开关读", 2),
    5: ("亮度读", 0),
    6: ("色温读", 4),
}
MODULE_SLOT_SIZE = 80
LIGHTS_PER_MODULE = 64
GROUPS_PER_MODULE = 16


@dataclass
class MainLayout:
    main: int
    max_sub: int
    modules: int
    light_slots: int
    group_slots: int
    light_offset: int
    group_offset: int
    module_offset: int


@dataclass
class Entry:
    address: str
    original_name: str
    generated_name: str
    final_name: str
    main: int
    middle: int
    sub: int
    module_in_main: int
    module_global: int
    object_type: str
    object_no: int
    func_name: str
    func_no: int


@dataclass
class SessionData:
    xml_bytes: bytes
    entries: List[Entry]
    summary: Dict[str, object]


SESSIONS: Dict[str, SessionData] = {}


def parse_address(address: str) -> Tuple[int, int, int]:
    parts = address.split("/")
    if len(parts) != 3:
        raise ValueError(f"Invalid group address: {address}")
    return int(parts[0]), int(parts[1]), int(parts[2])


def build_layout(main_to_max_sub: Dict[int, int]) -> Dict[int, MainLayout]:
    layouts: Dict[int, MainLayout] = {}
    light_offset = 0
    group_offset = 0
    module_offset = 0

    for main in sorted(main_to_max_sub):
        max_sub = main_to_max_sub[main]
        modules = max(1, math.ceil(max_sub / MODULE_SLOT_SIZE))
        full_modules, rem = divmod(max_sub, MODULE_SLOT_SIZE)
        light_slots = full_modules * LIGHTS_PER_MODULE + min(rem, LIGHTS_PER_MODULE)
        group_slots = full_modules * GROUPS_PER_MODULE + max(0, rem - LIGHTS_PER_MODULE)

        layouts[main] = MainLayout(
            main=main,
            max_sub=max_sub,
            modules=modules,
            light_slots=light_slots,
            group_slots=group_slots,
            light_offset=light_offset,
            group_offset=group_offset,
            module_offset=module_offset,
        )

        light_offset += light_slots
        group_offset += group_slots
        module_offset += modules

    return layouts


def generate_entries_from_xml(xml_bytes: bytes) -> Tuple[List[Entry], Dict[str, object]]:
    root = ET.fromstring(xml_bytes)
    candidates: List[Tuple[int, int, int, ET.Element]] = []
    main_to_max_sub: Dict[int, int] = {}

    for ga in root.findall(".//k:GroupAddress", NS):
        address = ga.get("Address", "")
        if not address:
            continue
        try:
            main, middle, sub = parse_address(address)
        except ValueError:
            continue
        if middle not in FUNCTIONS_BY_MIDDLE:
            continue
        if sub < 1:
            continue

        candidates.append((main, middle, sub, ga))
        prev = main_to_max_sub.get(main)
        if prev is None or sub > prev:
            main_to_max_sub[main] = sub

    layouts = build_layout(main_to_max_sub)

    entries: List[Entry] = []
    for main, middle, sub, ga in sorted(candidates, key=lambda x: (x[0], x[1], x[2])):
        layout = layouts[main]
        func_name, func_no = FUNCTIONS_BY_MIDDLE[middle]

        module_in_main = (sub - 1) // MODULE_SLOT_SIZE + 1
        module_global = layout.module_offset + module_in_main
        local_in_module = (sub - 1) % MODULE_SLOT_SIZE + 1

        if local_in_module <= LIGHTS_PER_MODULE:
            object_type = "灯"
            object_no = (
                layout.light_offset
                + (module_in_main - 1) * LIGHTS_PER_MODULE
                + local_in_module
            )
        else:
            object_type = "组"
            object_no = (
                layout.group_offset
                + (module_in_main - 1) * GROUPS_PER_MODULE
                + (local_in_module - LIGHTS_PER_MODULE)
            )

        base = f"{object_type}{object_no}"
        generated_name = f"{base} {DEVICE_NO} {func_no}"
        entries.append(
            Entry(
                address=ga.get("Address", ""),
                original_name=ga.get("Name", ""),
                generated_name=generated_name,
                final_name=generated_name,
                main=main,
                middle=middle,
                sub=sub,
                module_in_main=module_in_main,
                module_global=module_global,
                object_type=object_type,
                object_no=object_no,
                func_name=func_name,
                func_no=func_no,
            )
        )

    lights = sorted({e.object_no for e in entries if e.object_type == "灯"})
    groups = sorted({e.object_no for e in entries if e.object_type == "组"})

    summary = {
        "main_count": len(layouts),
        "module_count": sum(l.modules for l in layouts.values()),
        "light_count": len(lights),
        "group_count": len(groups),
        "converted_count": len(entries),
        "mains": [
            {
                "main": l.main,
                "max_sub": l.max_sub,
                "modules": l.modules,
                "light_slots": l.light_slots,
                "group_slots": l.group_slots,
            }
            for l in layouts.values()
        ],
    }

    return entries, summary


def apply_names_to_xml(xml_bytes: bytes, names_by_address: Dict[str, str]) -> bytes:
    tree = ET.ElementTree(ET.fromstring(xml_bytes))
    root = tree.getroot()

    for ga in root.findall(".//k:GroupAddress", NS):
        addr = ga.get("Address", "")
        new_name = names_by_address.get(addr)
        if new_name:
            ga.set("Name", new_name)

    out = io.BytesIO()
    tree.write(out, encoding="utf-8", xml_declaration=True)
    return out.getvalue()


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


def build_xlsx(entries: List[Entry], names_by_address: Dict[str, str]) -> bytes:
    headers = [
        "Address",
        "FinalName",
        "GeneratedName",
        "OriginalName",
        "Main",
        "Middle",
        "Sub",
        "ModuleInMain",
        "ModuleGlobal",
        "ObjectType",
        "ObjectNo",
        "DeviceNo",
        "FunctionNo",
        "FunctionName",
    ]

    rows = [headers]
    for e in entries:
        final_name = names_by_address.get(e.address, e.generated_name)
        rows.append(
            [
                e.address,
                final_name,
                e.generated_name,
                e.original_name,
                str(e.main),
                str(e.middle),
                str(e.sub),
                str(e.module_in_main),
                str(e.module_global),
                e.object_type,
                str(e.object_no),
                str(DEVICE_NO),
                str(e.func_no),
                e.func_name,
            ]
        )

    sheet_rows: List[str] = []
    for r_idx, row in enumerate(rows, start=1):
        cells = "".join(inline_cell(r_idx, c_idx, val) for c_idx, val in enumerate(row, start=1))
        sheet_rows.append(f'<row r="{r_idx}">{cells}</row>')

    dimension = f"A1:{excel_col(len(headers))}{len(rows)}"

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
        f"<dc:creator>{PROJECT_NAME_EN}</dc:creator>"
        f"<cp:lastModifiedBy>{PROJECT_NAME_EN}</cp:lastModifiedBy>"
        "</cp:coreProperties>"
    )

    app_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        f"<Application>{PROJECT_NAME_EN}</Application>"
        "</Properties>"
    )

    out = io.BytesIO()
    with ZipFile(out, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("docProps/core.xml", core_xml)
        zf.writestr("docProps/app.xml", app_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        zf.writestr("xl/worksheets/sheet1.xml", worksheet_xml)
    return out.getvalue()


def build_default_name_map(entries: List[Entry]) -> Dict[str, str]:
    return {e.address: e.generated_name for e in entries}


class Handler(BaseHTTPRequestHandler):
    def _write_bytes(self, status: int, content_type: str, data: bytes, extra_headers: Dict[str, str] | None = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        if extra_headers:
            for k, v in extra_headers.items():
                self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def _write_json(self, status: int, obj: Dict[str, object]) -> None:
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self._write_bytes(status, "application/json; charset=utf-8", data)

    def _read_json_body(self) -> Dict[str, object]:
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            return {}
        raw = self.rfile.read(length)
        if not raw:
            return {}
        return json.loads(raw.decode("utf-8"))

    def _read_multipart_file(self, field_name: str) -> Tuple[bytes | None, str | None]:
        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            return None, "Use multipart/form-data upload"

        m = re.search(r'boundary="?([^\";]+)"?', content_type)
        if not m:
            return None, "Missing multipart boundary"
        boundary = m.group(1).encode("utf-8")

        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            return None, "Invalid Content-Length"
        body = self.rfile.read(length)
        if not body:
            return None, "Uploaded body is empty"

        marker = b"--" + boundary
        for part in body.split(marker):
            part = part.strip()
            if not part or part == b"--":
                continue
            head, sep, data = part.partition(b"\r\n\r\n")
            if not sep:
                continue
            headers_text = head.decode("utf-8", errors="ignore")
            if f'name="{field_name}"' not in headers_text:
                continue
            if data.endswith(b"\r\n"):
                data = data[:-2]
            if data.endswith(b"--"):
                data = data[:-2]
            if not data:
                return None, f"{field_name} is empty"
            return data, None

        return None, f"{field_name} is required"

    def do_GET(self) -> None:
        if self.path == "/":
            self._write_bytes(HTTPStatus.OK, "text/html; charset=utf-8", INDEX_HTML.encode("utf-8"))
            return
        self._write_json(HTTPStatus.NOT_FOUND, {"error": "Not found"})

    def do_POST(self) -> None:
        if self.path == "/api/parse":
            self.handle_parse()
            return
        if self.path == "/api/export/xml":
            self.handle_export_xml()
            return
        if self.path == "/api/export/xlsx":
            self.handle_export_xlsx()
            return
        self._write_json(HTTPStatus.NOT_FOUND, {"error": "Not found"})

    def handle_parse(self) -> None:
        xml_bytes, err = self._read_multipart_file("xml_file")
        if err:
            self._write_json(HTTPStatus.BAD_REQUEST, {"error": err})
            return
        assert xml_bytes is not None

        try:
            entries, summary = generate_entries_from_xml(xml_bytes)
        except ET.ParseError as exc:
            self._write_json(HTTPStatus.BAD_REQUEST, {"error": f"XML parse failed: {exc}"})
            return
        except Exception as exc:  # noqa: BLE001
            self._write_json(HTTPStatus.BAD_REQUEST, {"error": f"Parse failed: {exc}"})
            return

        session_id = uuid.uuid4().hex
        SESSIONS[session_id] = SessionData(xml_bytes=xml_bytes, entries=entries, summary=summary)

        self._write_json(
            HTTPStatus.OK,
            {
                "session_id": session_id,
                "summary": summary,
                "entries": [asdict(e) for e in entries],
            },
        )

    def _resolve_export_payload(self) -> Tuple[SessionData | None, Dict[str, str] | None, str | None]:
        payload = self._read_json_body()
        session_id = str(payload.get("session_id", ""))
        if not session_id:
            return None, None, "session_id is required"

        session = SESSIONS.get(session_id)
        if session is None:
            return None, None, "Session expired or not found"

        names_payload = payload.get("names", {})
        names_by_address: Dict[str, str] = build_default_name_map(session.entries)
        if isinstance(names_payload, dict):
            for k, v in names_payload.items():
                if not isinstance(k, str):
                    continue
                if not isinstance(v, str):
                    continue
                if v.strip() == "":
                    continue
                names_by_address[k] = v.strip()

        return session, names_by_address, None

    def handle_export_xml(self) -> None:
        session, names_by_address, err = self._resolve_export_payload()
        if err:
            self._write_json(HTTPStatus.BAD_REQUEST, {"error": err})
            return
        assert session is not None and names_by_address is not None

        xml_bytes = apply_names_to_xml(session.xml_bytes, names_by_address)
        self._write_bytes(
            HTTPStatus.OK,
            "application/xml",
            xml_bytes,
            {
                "Content-Disposition": 'attachment; filename="knx_converted.xml"',
            },
        )

    def handle_export_xlsx(self) -> None:
        session, names_by_address, err = self._resolve_export_payload()
        if err:
            self._write_json(HTTPStatus.BAD_REQUEST, {"error": err})
            return
        assert session is not None and names_by_address is not None

        xlsx_bytes = build_xlsx(session.entries, names_by_address)
        self._write_bytes(
            HTTPStatus.OK,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            xlsx_bytes,
            {
                "Content-Disposition": 'attachment; filename="knx_converted.xlsx"',
            },
        )


INDEX_HTML = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>KNX 群组地址转换工作台</title>
  <style>
    :root {
      --bg: #f3f5f4;
      --panel: #ffffff;
      --line: #d9dedb;
      --text: #1f2a24;
      --accent: #1b7f5e;
      --accent-soft: #e8f4ef;
      --warn: #b94a48;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "PingFang SC", "Microsoft YaHei", sans-serif;
      color: var(--text);
      background:
        radial-gradient(circle at 8% 5%, #ffffff 0%, #f3f5f4 38%),
        linear-gradient(160deg, #eef3f1 0%, #f6f7f6 60%, #eef3f1 100%);
    }
    .wrap {
      max-width: 1400px;
      margin: 0 auto;
      padding: 16px;
      display: grid;
      gap: 12px;
    }
    .card {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 12px;
    }
    h1 {
      margin: 0;
      font-size: 22px;
      letter-spacing: 0.5px;
    }
    .toolbar {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      align-items: center;
    }
    .muted { color: #607068; font-size: 13px; }
    button, .btn {
      border: 1px solid var(--line);
      background: #fff;
      color: var(--text);
      border-radius: 8px;
      padding: 8px 12px;
      cursor: pointer;
      font-size: 13px;
    }
    button:hover, .btn:hover { border-color: var(--accent); }
    .primary {
      border-color: var(--accent);
      background: var(--accent-soft);
      color: #124f3b;
      font-weight: 600;
    }
    input[type="text"], input[type="number"], input[type="search"] {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 8px;
      font-size: 13px;
      min-width: 120px;
      background: #fff;
    }
    .summary {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
      gap: 8px;
    }
    .kpi {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 8px;
      background: #fafcfa;
    }
    .kpi .label { font-size: 12px; color: #63726a; }
    .kpi .value { font-size: 19px; font-weight: 700; }
    .table-wrap {
      border: 1px solid var(--line);
      border-radius: 10px;
      overflow: auto;
      max-height: 70vh;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
      min-width: 1100px;
    }
    th, td {
      border-bottom: 1px solid #e8ecea;
      padding: 6px;
      text-align: left;
      vertical-align: middle;
      white-space: nowrap;
    }
    th {
      position: sticky;
      top: 0;
      z-index: 2;
      background: #f7faf8;
      border-bottom: 1px solid #d9dedb;
    }
    td input[type="text"] {
      width: 280px;
      font-size: 12px;
      padding: 5px;
    }
    .row-info {
      display: flex;
      gap: 8px;
      align-items: center;
      flex-wrap: wrap;
      margin-top: 8px;
    }
    .error { color: var(--warn); font-weight: 600; }
    .ok { color: #206a4f; font-weight: 600; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <h1>KNX 群组地址转换工作台</h1>
      <div class="muted">上传 ETS XML -> 自动编号 -> 网页修改 -> 导出 XML 和 Excel</div>
    </div>

    <div class="card">
      <div class="toolbar">
        <input id="xmlFile" type="file" accept=".xml" />
        <button id="parseBtn" class="primary">上传并解析 XML</button>
        <button id="downloadXmlBtn" disabled>下载 XML</button>
        <button id="downloadXlsxBtn" disabled>下载 Excel</button>
        <button id="downloadTemplateBtn" disabled>下载批量模板(CSV)</button>
        <label class="btn" for="uploadTemplateInput">上传模板(CSV)</label>
        <input id="uploadTemplateInput" type="file" accept=".csv" style="display:none" />
      </div>
      <div class="row-info">
        <input id="searchInput" type="search" placeholder="按地址/名称筛选" />
        <label>每页
          <input id="pageSizeInput" type="number" min="50" step="50" value="200" style="width: 90px;" />
          条
        </label>
        <button id="prevPageBtn" disabled>上一页</button>
        <button id="nextPageBtn" disabled>下一页</button>
        <span id="pageInfo" class="muted"></span>
      </div>
      <div id="status" class="muted" style="margin-top:8px"></div>
    </div>

    <div class="card summary" id="summary"></div>

    <div class="card">
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>地址</th>
              <th>原始 Name</th>
              <th>自动 Name</th>
              <th>当前 Name(可编辑)</th>
              <th>对象</th>
              <th>对象编号</th>
              <th>功能</th>
              <th>模块</th>
              <th>主/中/子</th>
            </tr>
          </thead>
          <tbody id="tableBody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <script>
    let sessionId = "";
    let allEntries = [];
    let filteredIndexes = [];
    let currentPage = 1;

    const statusEl = document.getElementById("status");
    const summaryEl = document.getElementById("summary");
    const tableBodyEl = document.getElementById("tableBody");
    const searchInput = document.getElementById("searchInput");
    const pageSizeInput = document.getElementById("pageSizeInput");
    const pageInfo = document.getElementById("pageInfo");

    function setStatus(msg, ok = true) {
      statusEl.className = ok ? "ok" : "error";
      statusEl.textContent = msg;
    }

    function escapeHtml(text) {
      return String(text)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
    }

    function parseBaseName(finalName, funcNo) {
      const suffix = ` 9 ${funcNo}`;
      if (finalName.endsWith(suffix)) {
        return finalName.slice(0, -suffix.length);
      }
      return finalName;
    }

    function downloadBlob(blob, filename) {
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    }

    function updateButtons(enabled) {
      document.getElementById("downloadXmlBtn").disabled = !enabled;
      document.getElementById("downloadXlsxBtn").disabled = !enabled;
      document.getElementById("downloadTemplateBtn").disabled = !enabled;
      document.getElementById("prevPageBtn").disabled = !enabled;
      document.getElementById("nextPageBtn").disabled = !enabled;
    }

    function renderSummary(summary) {
      if (!summary) {
        summaryEl.innerHTML = "";
        return;
      }
      const items = [
        ["主群组数量", summary.main_count],
        ["检测模块数", summary.module_count],
        ["灯数量", summary.light_count],
        ["组数量", summary.group_count],
        ["转换条目", summary.converted_count],
      ];
      summaryEl.innerHTML = items.map(([label, value]) => `
        <div class="kpi">
          <div class="label">${escapeHtml(label)}</div>
          <div class="value">${escapeHtml(value)}</div>
        </div>
      `).join("");
    }

    function recalcFilter() {
      const q = searchInput.value.trim().toLowerCase();
      filteredIndexes = [];
      for (let i = 0; i < allEntries.length; i += 1) {
        const e = allEntries[i];
        if (!q) {
          filteredIndexes.push(i);
          continue;
        }
        const hit = (
          e.address.toLowerCase().includes(q) ||
          e.original_name.toLowerCase().includes(q) ||
          e.generated_name.toLowerCase().includes(q) ||
          e.final_name.toLowerCase().includes(q) ||
          `${e.object_type}${e.object_no}`.toLowerCase().includes(q)
        );
        if (hit) filteredIndexes.push(i);
      }
      currentPage = 1;
      renderTable();
    }

    function getPageSize() {
      const n = Number(pageSizeInput.value || 200);
      if (!Number.isFinite(n) || n <= 0) return 200;
      return Math.max(50, Math.floor(n));
    }

    function renderTable() {
      const pageSize = getPageSize();
      const total = filteredIndexes.length;
      const totalPages = Math.max(1, Math.ceil(total / pageSize));
      if (currentPage > totalPages) currentPage = totalPages;
      if (currentPage < 1) currentPage = 1;

      const start = (currentPage - 1) * pageSize;
      const end = Math.min(total, start + pageSize);

      const rows = [];
      for (let k = start; k < end; k += 1) {
        const idx = filteredIndexes[k];
        const e = allEntries[idx];
        rows.push(`
          <tr>
            <td>${escapeHtml(e.address)}</td>
            <td title="${escapeHtml(e.original_name)}">${escapeHtml(e.original_name)}</td>
            <td title="${escapeHtml(e.generated_name)}">${escapeHtml(e.generated_name)}</td>
            <td>
              <input type="text" data-idx="${idx}" class="name-input" value="${escapeHtml(e.final_name)}" />
            </td>
            <td>${escapeHtml(e.object_type)}</td>
            <td>${escapeHtml(e.object_no)}</td>
            <td>${escapeHtml(e.func_name)}(${escapeHtml(e.func_no)})</td>
            <td>${escapeHtml(e.module_global)} / ${escapeHtml(e.module_in_main)}</td>
            <td>${escapeHtml(e.main)}/${escapeHtml(e.middle)}/${escapeHtml(e.sub)}</td>
          </tr>
        `);
      }
      tableBodyEl.innerHTML = rows.join("");

      pageInfo.textContent = `第 ${currentPage}/${totalPages} 页，显示 ${start + 1}-${end}，共 ${total} 条`;
      document.getElementById("prevPageBtn").disabled = currentPage <= 1;
      document.getElementById("nextPageBtn").disabled = currentPage >= totalPages;

      document.querySelectorAll(".name-input").forEach((input) => {
        input.addEventListener("input", (ev) => {
          const idx = Number(ev.target.dataset.idx);
          if (Number.isFinite(idx) && allEntries[idx]) {
            allEntries[idx].final_name = ev.target.value;
          }
        });
      });
    }

    function buildNamesMap() {
      const map = {};
      for (const e of allEntries) {
        map[e.address] = e.final_name;
      }
      return map;
    }

    function buildTemplateCsv() {
      const uniq = new Map();
      for (const e of allEntries) {
        const key = `${e.object_type}|${e.object_no}`;
        if (!uniq.has(key)) {
          uniq.set(key, {
            object_type: e.object_type,
            object_no: e.object_no,
            base_name: parseBaseName(e.final_name, e.func_no),
          });
        }
      }
      const rows = ["object_type,object_no,base_name"];
      for (const item of uniq.values()) {
        const base = String(item.base_name).replaceAll('"', '""');
        rows.push(`${item.object_type},${item.object_no},"${base}"`);
      }
      return rows.join("\n");
    }

    function parseCsvLine(line) {
      const out = [];
      let cur = "";
      let inQuote = false;
      for (let i = 0; i < line.length; i += 1) {
        const ch = line[i];
        if (ch === '"') {
          if (inQuote && line[i + 1] === '"') {
            cur += '"';
            i += 1;
          } else {
            inQuote = !inQuote;
          }
        } else if (ch === ',' && !inQuote) {
          out.push(cur);
          cur = "";
        } else {
          cur += ch;
        }
      }
      out.push(cur);
      return out;
    }

    function applyTemplateCsv(text) {
      const lines = text.replaceAll("\r\n", "\n").replaceAll("\r", "\n").split("\n").filter(Boolean);
      if (lines.length < 2) {
        throw new Error("模板内容为空");
      }
      const header = parseCsvLine(lines[0]).map((x) => x.trim());
      const idxType = header.indexOf("object_type");
      const idxNo = header.indexOf("object_no");
      const idxBase = header.indexOf("base_name");
      if (idxType < 0 || idxNo < 0 || idxBase < 0) {
        throw new Error("模板字段必须包含 object_type, object_no, base_name");
      }

      const baseMap = new Map();
      for (let i = 1; i < lines.length; i += 1) {
        const cols = parseCsvLine(lines[i]);
        if (!cols.length) continue;
        const t = (cols[idxType] || "").trim();
        const n = Number((cols[idxNo] || "").trim());
        const b = (cols[idxBase] || "").trim();
        if (!t || !Number.isFinite(n) || !b) continue;
        baseMap.set(`${t}|${n}`, b);
      }

      let changed = 0;
      for (const e of allEntries) {
        const key = `${e.object_type}|${e.object_no}`;
        const base = baseMap.get(key);
        if (!base) continue;
        e.final_name = `${base} 9 ${e.func_no}`;
        changed += 1;
      }
      renderTable();
      setStatus(`模板已应用，更新 ${changed} 条名称`, true);
    }

    async function parseXml() {
      const fileInput = document.getElementById("xmlFile");
      if (!fileInput.files || !fileInput.files[0]) {
        setStatus("请先选择 XML 文件", false);
        return;
      }
      const fd = new FormData();
      fd.append("xml_file", fileInput.files[0]);

      setStatus("正在解析 XML...", true);
      const resp = await fetch("/api/parse", { method: "POST", body: fd });
      const data = await resp.json();
      if (!resp.ok) {
        setStatus(data.error || "解析失败", false);
        return;
      }

      sessionId = data.session_id;
      allEntries = data.entries || [];
      renderSummary(data.summary || null);
      recalcFilter();
      updateButtons(true);
      setStatus(`解析成功，已加载 ${allEntries.length} 条`, true);
    }

    async function exportFile(kind) {
      if (!sessionId) {
        setStatus("请先上传并解析 XML", false);
        return;
      }
      const payload = {
        session_id: sessionId,
        names: buildNamesMap(),
      };
      const path = kind === "xml" ? "/api/export/xml" : "/api/export/xlsx";
      const filename = kind === "xml" ? "knx_converted.xml" : "knx_converted.xlsx";

      setStatus(`正在生成 ${filename} ...`, true);
      const resp = await fetch(path, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (!resp.ok) {
        let err = "导出失败";
        try {
          const e = await resp.json();
          err = e.error || err;
        } catch (_ignore) {}
        setStatus(err, false);
        return;
      }
      const blob = await resp.blob();
      downloadBlob(blob, filename);
      setStatus(`${filename} 已下载`, true);
    }

    document.getElementById("parseBtn").addEventListener("click", () => {
      parseXml().catch((e) => setStatus(`解析异常: ${e.message}`, false));
    });
    document.getElementById("downloadXmlBtn").addEventListener("click", () => {
      exportFile("xml").catch((e) => setStatus(`导出异常: ${e.message}`, false));
    });
    document.getElementById("downloadXlsxBtn").addEventListener("click", () => {
      exportFile("xlsx").catch((e) => setStatus(`导出异常: ${e.message}`, false));
    });

    document.getElementById("downloadTemplateBtn").addEventListener("click", () => {
      const text = buildTemplateCsv();
      downloadBlob(new Blob([text], { type: "text/csv;charset=utf-8" }), "batch_template.csv");
      setStatus("模板已下载", true);
    });

    document.getElementById("uploadTemplateInput").addEventListener("change", async (ev) => {
      const f = ev.target.files && ev.target.files[0];
      if (!f) return;
      try {
        const text = await f.text();
        applyTemplateCsv(text);
      } catch (e) {
        setStatus(`模板导入失败: ${e.message}`, false);
      } finally {
        ev.target.value = "";
      }
    });

    searchInput.addEventListener("input", recalcFilter);
    pageSizeInput.addEventListener("change", renderTable);
    document.getElementById("prevPageBtn").addEventListener("click", () => {
      currentPage -= 1;
      renderTable();
    });
    document.getElementById("nextPageBtn").addEventListener("click", () => {
      currentPage += 1;
      renderTable();
    });

    updateButtons(false);
    setStatus("请先上传 XML", true);
  </script>
</body>
</html>
"""


def launch_url(host: str, port: int) -> str:
    target_host = host
    if host in {"0.0.0.0", "::", ""}:
        target_host = "127.0.0.1"
    return f"http://{target_host}:{port}"


def open_browser_async(url: str, delay: float = 0.8) -> None:
    def _open() -> None:
        try:
            webbrowser.open(url, new=2)
        except Exception as exc:  # noqa: BLE001
            print(f"Auto-open browser failed: {exc}")

    timer = threading.Timer(delay, _open)
    timer.daemon = True
    timer.start()


def run_server(host: str, port: int, auto_open_browser: bool = True) -> None:
    url = launch_url(host, port)
    server = ThreadingHTTPServer((host, port), Handler)
    print(f"{PROJECT_NAME_EN} running at: {url}")
    if auto_open_browser:
        print(f"Opening browser: {url}")
        open_browser_async(url)
    print("Press Ctrl+C to stop")
    server.serve_forever()


def main() -> None:
    parser = argparse.ArgumentParser(description="KNX XML -> web rename tool")
    parser.add_argument("--host", default="127.0.0.1", help="Bind host (default: 127.0.0.1)")
    parser.add_argument("--port", type=int, default=8765, help="Bind port (default: 8765)")
    parser.add_argument("--no-browser", action="store_true", help="Do not auto-open web browser")
    args = parser.parse_args()

    run_server(args.host, args.port, auto_open_browser=not args.no_browser)


if __name__ == "__main__":
    main()
