# 架构设计文档

## 1. 设计目标

针对**"用户有 Excel 但没有 Python"**的约束：
- **开发端**：Python + pywin32 + VS Code 调试，工程化开发
- **用户端**：PyInstaller 打包为单文件 exe，双击运行，无需 Python
- **格式保真**：通过 Excel COM `SaveAs` 实现 100% 格式保留

## 2. 技术选型

| 组件 | 选型 | 理由 |
|------|------|------|
| 脚本语言 | Python 3.9+ | 开发效率高，生态丰富，VS Code 支持极佳 |
| XLS 读取/写入 | Excel COM (`pywin32`) | 用户已有 Excel，格式保留最完整 |
| 打包工具 | PyInstaller | 将 Python 脚本+解释器+依赖打包为单文件 exe |
| 测试框架 | pytest | 标准 Python 测试框架 |
| 静态检查 | flake8 | 代码风格与语法检查 |
| CI/CD | GitHub Actions | 自动测试 + 自动构建 Windows exe + Release 发布 |
| IDE | VS Code + Python 扩展 | 断点调试 COM 对象、变量监视、终端集成 |

## 3. 模块职责

```
main.py (根入口，打包后的 exe 入口点)
    | argparse 解析参数
    | sys.path 注入 src/
src/xls2xlsx/
    ├── scanner.py      # 文件系统扫描：递归查找 .xls，严格过滤
    ├── converter.py    # Excel COM 生命周期：启动、单文件转换、安全退出
    ├── cli.py          # 交互层：彩色输出、扫描报告、警告、确认、进度、汇总
    └── main.py         # 主流程编排
```

### scanner.py
- `scan_xls_files()`：使用 `pathlib.Path.glob` 递归匹配，严格过滤 `.xls` 后缀
- `format_size()`：辅助函数，字节转人类可读格式

### converter.py
- `start_excel_instance()`：`win32com.client.Dispatch("Excel.Application")`，设置静默模式
- `convert_single_xls()`：`Workbooks.Open()` -> `SaveAs(FileFormat=51)` -> `Close(SaveChanges=False)`
- `stop_excel_instance()`：`Quit()` + `del` + `gc.collect()`，防止进程残留

### cli.py
- 纯 ANSI 转义码彩色输出，不依赖外部库
- 四层确认流程：扫描报告 -> 备份警告 -> 模式选择 -> 最终确认

## 4. 数据流

```
[文件系统]
    | (scanner.py: Path.glob)
[.xls 文件列表 + 统计]
    | (cli.py: 扫描报告 + 用户确认)
[待转换队列]
    | (converter.py: start_excel_instance)
Excel.Application (后台静默)
    | 逐文件
Workbooks.Open(.xls) -> SaveAs(.xlsx, 51) -> Close
    | (cli.py: 进度 + 汇总)
[转换报告]
    | (converter.py: stop_excel_instance)
Excel 进程安全退出
```

## 5. 打包说明

PyInstaller 将以下组件打包为单文件 exe：
- Python 解释器（嵌入式）
- pywin32 库及 COM 桥接层
- 项目源码（scanner.py, converter.py, cli.py, main.py）
- 运行时依赖

最终用户只需：
1. 下载 `xls2xlsx.exe`
2. 放到目标文件夹或任意位置
3. 双击运行（或在 CMD/PowerShell 中执行）

## 6. 异常处理

| 异常场景 | 处理策略 |
|----------|----------|
| Excel 未安装 | 启动时捕获，提示用户安装 Excel |
| 源文件无法打开 | 单文件标记 Failed，继续批量任务 |
| 目标文件已存在 | 标记 Skipped，不覆盖 |
| Excel 进程残留 | `finally` 块强制 `stop_excel_instance()` |
| 用户取消确认 | 优雅退出，返回 code 0，Excel 未启动 |
| 删除原文件失败 | 记录 warning，不影响整体汇总 |
