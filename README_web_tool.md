# KNX 设备9 网页转换工具

## 1. 启动

```bash
python3 "/Users/han/Documents/New project/knx_web_tool.py" --host 127.0.0.1 --port 8765
```

浏览器打开：

- <http://127.0.0.1:8765>

## 2. 使用流程

1. 上传 ETS 导出的 XML 文件。
2. 工具自动按设备9规则生成 Name。
3. 在网页表格中查看/手动修改 `群组地址 + Name`。
4. 可下载 `模板(CSV)` 批量修改灯名/组名，再上传模板应用。
5. 下载导出的 `XML` 和 `Excel(xlsx)`。

## 3. 当前自动编号规则

- 功能映射（中群组）：
  - `1 -> 开关写 -> 功能号3`
  - `2 -> 亮度写 -> 功能号1`
  - `3 -> 色温写 -> 功能号5`
  - `4 -> 开关读 -> 功能号2`
  - `5 -> 亮度读 -> 功能号0`
  - `6 -> 色温读 -> 功能号4`
- 自动检测每个主群组的 `max_sub` 和模块数（`每80个子地址=1个模块`）。
- 每个模块固定拆分：
  - `sub 1~64 -> 灯`
  - `sub 65~80 -> 组`
- 例如 `max_sub=160` 时（2个模块）：
  - `1~64 -> 灯1~64`
  - `65~80 -> 组1~16`
  - `81~144 -> 灯65~128`
  - `145~160 -> 组17~32`
- 灯号/组号在主群组之间连续递增（按“实际灯位/组位数量”连续编号）。

## 4. 批量模板格式

模板 CSV 字段：

- `object_type`
- `object_no`
- `base_name`

示例：

```csv
object_type,object_no,base_name
灯,1,客厅主灯
灯,2,客厅辅灯
组,1,场景组A
```

应用后 Name 自动拼接为：`base_name 9 功能号`。

## 5. Mac/Windows

- 只要有 Python 3 都可运行。
- Windows 启动命令：

```powershell
python "C:\path\to\knx_web_tool.py" --host 127.0.0.1 --port 8765
```

## 6. 可执行文件封装（可选）

如果你希望无需 Python 运行，可用项目自带脚本打包：

- macOS:

```bash
bash "/Users/han/Documents/New project/scripts/build_macos.sh"
```

- Windows (PowerShell):

```powershell
powershell -ExecutionPolicy Bypass -File "C:\path\to\scripts\build_windows.ps1"
```

输出文件：

- `knx-web-tool-macos.zip`
- `knx-web-tool-windows.zip`

## 7. GitHub 自动构建与发布

仓库已包含工作流：

- `/Users/han/Documents/New project/.github/workflows/build-release.yml`

触发方式：

1. 推送标签（例如 `v1.0.0`）到 GitHub，会自动构建 Win+Mac 并发布 Release。
2. 在 GitHub Actions 页面手动运行 `Build And Release`，输入标签名发布。
