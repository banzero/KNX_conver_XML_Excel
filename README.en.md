# KNX Group Address Studio

[中文说明（默认）](./README.md)

`KNX Group Address Studio` is a local web tool that converts ETS-exported KNX group-address XML into MaiLian-compatible naming (Device No. 9: KNX DALI Tunable White Dimmer), with online editing and XML/XLSX export.

## What this project does

- Upload and parse ETS XML group addresses.
- Auto-generate names using Device 9 mapping (`object 9 functionNo`).
- Review all addresses and names in a web table.
- Edit names row-by-row.
- Batch rename via CSV template import.
- Export updated XML and Excel (`.xlsx`).
- Auto-open browser page when the server starts.

## Quick start

### 1) Run the tool (browser opens automatically)

```bash
python3 knx_web_tool.py
```

Default URL: `http://127.0.0.1:8765`

Disable browser auto-open:

```bash
python3 knx_web_tool.py --no-browser
```

### 2) Workflow in browser

1. Upload XML and parse.
2. Review generated names.
3. Edit any row if needed.
4. For batch rename:
   - Download CSV template.
   - Fill `object_type, object_no, base_name`.
   - Upload CSV template.
5. Export XML / Excel.

## Current mapping rules (Device 9)

### Function mapping (middle group)

- `1 -> Switch Write -> 3`
- `2 -> Brightness Write -> 1`
- `3 -> Color Temp Write -> 5`
- `4 -> Switch Read -> 2`
- `5 -> Brightness Read -> 0`
- `6 -> Color Temp Read -> 4`

### Module address split

- Each `80` sub-addresses form one module.
- In each module:
  - `1~64` are lights
  - `65~80` are groups
- Light and group numbering is continuous across main groups.

## Batch template example

```csv
object_type,object_no,base_name
灯,1,LivingRoom_Main
灯,2,LivingRoom_Secondary
组,1,SceneGroup_A
```

Final name format: `base_name 9 functionNo`.

## Packaging (optional)

### macOS

```bash
bash scripts/build_macos.sh
```

### Windows (PowerShell)

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_windows.ps1
```

Output files:

- `knx-web-tool-macos.zip`
- `knx-web-tool-windows.zip`

## GitHub release automation

Workflow file: `.github/workflows/build-release.yml`

- Push a tag (for example `v0.0.2`) to auto-build Win + macOS packages and publish a Release.
- Or run the workflow manually from GitHub Actions.

## Project structure

- `knx_web_tool.py`: web app entrypoint.
- `knx_mylink_converter.py`: CLI converter.
- `scripts/`: local packaging scripts.
- `.github/workflows/`: CI build/release workflows.
