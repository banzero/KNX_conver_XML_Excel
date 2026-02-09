# KNX 群组地址转换工作台

[English Documentation](./README.en.md)

`KNX 群组地址转换工作台` 是一个本地网页工具，用于把 ETS 导出的 KNX 群组地址 XML，按迈联主机（设备编号 9：KNX DALI 双色温调光灯）规则自动映射、批量改名并导出。

## 这个项目能做什么

- 上传 ETS 导出的 XML 文件并自动解析群组地址。
- 按设备 9 规则自动生成名称（`对象名 9 功能号`）。
- 在网页中查看所有群组地址和 Name，支持逐条手工修改。
- 下载 CSV 模板，批量修改灯名/组名后回传应用。
- 导出修改后的 XML 和 Excel（`.xlsx`）。
- 启动服务后自动打开浏览器页面。

## 适用对象

- 需要把 KNX 群组地址映射到迈联导入格式的工程人员。
- 需要快速校对大量灯控对象名称的集成人员。

## 快速开始

### 1) 启动（自动打开浏览器）

```bash
python3 knx_web_tool.py
```

默认地址：`http://127.0.0.1:8765`

如果你不想自动打开浏览器：

```bash
python3 knx_web_tool.py --no-browser
```

### 2) 网页使用流程

1. 点击“上传并解析 XML”。
2. 检查自动生成的 Name。
3. 需要调整时，可直接在表格里改。
4. 如需批量改名：
   - 先下载模板 CSV。
   - 填写 `object_type, object_no, base_name`。
   - 上传模板并应用。
5. 下载导出的 XML / Excel。

## 当前映射规则（设备编号 9）

### 功能映射（中群组）

- `1 -> 开关写 -> 功能号 3`
- `2 -> 亮度写 -> 功能号 1`
- `3 -> 色温写 -> 功能号 5`
- `4 -> 开关读 -> 功能号 2`
- `5 -> 亮度读 -> 功能号 0`
- `6 -> 色温读 -> 功能号 4`

### 模块地址拆分

- 每 `80` 个子地址算 1 个模块。
- 每个模块固定：
  - `1~64` 为灯
  - `65~80` 为组
- 灯号和组号会在主群组间连续编号。

## 批量模板示例

```csv
object_type,object_no,base_name
灯,1,客厅主灯
灯,2,客厅辅灯
组,1,场景组A
```

应用后会自动拼成：`base_name 9 功能号`。

## 打包（可选）

### macOS

```bash
bash scripts/build_macos.sh
```

### Windows (PowerShell)

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_windows.ps1
```

输出文件：

- `knx-web-tool-macos.zip`
- `knx-web-tool-windows.zip`

## GitHub 自动发布

项目已内置工作流：`.github/workflows/build-release.yml`

- 推送标签（如 `v0.0.2`）后，会自动构建 Win + Mac 并发布 Release。
- 也可在 GitHub Actions 手动触发工作流。

## 目录说明

- `knx_web_tool.py`：网页工具主程序。
- `knx_mylink_converter.py`：命令行转换脚本。
- `scripts/`：本地打包脚本。
- `.github/workflows/`：CI 构建与发布脚本。

## 许可

如果你有内部交付规范，可在本项目补充 `LICENSE` 文件。
