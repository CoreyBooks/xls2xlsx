# VS Code 调试指南

## 环境要求

- Windows 10/11
- Python 3.9+
- VS Code + Python 扩展
- Microsoft Excel（任意版本）

## 快速开始

```bash
code xls2xlsx-converter-py
```

## 调试配置

仓库已内置 6 套配置（`.vscode/launch.json`）：

### 运行转换工具 (当前目录)
- 完整交互流程调试
- 观察点：`cli.py` 的 `confirm_backup()` 和 `choose_mode()`

### 运行转换工具 (指定测试目录)
- 带 `test_data/` 参数调试
- 观察点：`scanner.py` 的 `scan_xls_files()` 返回结果

### 运行转换工具 (强制删除模式)
- `--delete --force`，跳过交互
- 观察点：`converter.py` 的文件删除逻辑

### 调试扫描模块
- 独立调试 `scanner.py`，无需 Excel
- 自动创建临时测试数据

### 调试转换模块
- 独立调试 `converter.py`，需真实 `.xls` + Excel
- 观察点：
  - `start_excel_instance()` 后 `$excel` 属性
  - `convert_single_xls()` 中 `SaveAs` 调用
  - `stop_excel_instance()` 后任务管理器中 Excel 进程

### 运行 pytest
- 执行全部单元测试

## 关键断点

| 文件 | 函数 | 观察内容 |
|------|------|----------|
| `scanner.py` | `scan_xls_files()` | `result.files`, `result.count` |
| `converter.py` | `start_excel_instance()` | `excel.Visible`, `excel.DisplayAlerts` |
| `converter.py` | `convert_single_xls()` | `abs_src`, `abs_dst`, `wb.Sheets.Count` |
| `converter.py` | `stop_excel_instance()` | 任务管理器中 EXCEL.EXE 是否消失 |
| `cli.py` | `show_scan_report()` | 彩色输出与格式化结果 |

## 常见问题

### Q1: 调试时提示 ModuleNotFoundError: No module named 'win32com'
**解决**：确保虚拟环境已激活，且已安装 pywin32：
```bash
pip install pywin32
```

### Q2: 转换后 Excel 进程未退出
**解决**：在 `finally` 块中检查 `excel` 是否为 `None`。若仍残留：
```python
import os
os.system("taskkill /f /im excel.exe")
```

### Q3: 如何准备测试用的 .xls 文件？
在 Excel 中：
1. 新建工作簿
2. 输入数据
3. 另存为 -> Excel 97-2003 工作簿 (*.xls)
4. 放入 `test_data/` 目录

### Q4: PyInstaller 打包后 exe 无法运行
**解决**：确保在 Windows 环境下打包，且 pywin32 已安装：
```bash
pip install pyinstaller pywin32
pyinstaller --onefile --name xls2xlsx main.py
```
