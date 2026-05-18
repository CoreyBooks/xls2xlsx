# XLS -> XLSX 批量转换工具 (Python + Excel COM)

> **利用现有 Excel . 打包为单文件 exe . 批量递归 . 安全确认**

针对**已有 Microsoft Excel** 的用户群体设计，提供 **exe / Python 脚本 / VBS 脚本** 三种方案，覆盖不同部署环境：
- **exe**：单文件，无需安装任何东西，双击运行
- **py 脚本**：有 Python 环境的用户，可自定义扩展
- **VBS 脚本**：企业受限电脑（禁止外部 exe / 无 Python），仅需 Excel 即可运行

---

## 核心特性

| 特性 | 说明 |
|------|------|
| **三格式发布** | 提供 exe（无 Python）、py 脚本（需 pywin32）、VBS（仅需 Excel） |
| **格式 100% 保留** | 通过 Excel COM `SaveAs(FileFormat=51)` 另存为，与手动操作效果完全一致 |
| **批量递归** | 自动扫描当前目录及所有子文件夹中的 `.xls` 文件 |
| **安全确认** | 扫描统计 -> 备份警告 -> 模式选择 -> 最终确认，四层防护 |
| **工程化** | 模块化源码、pytest 测试、flake8 检查、GitHub Actions CI/CD |
| **VS Code 调试** | 内置调试配置，支持断点观察 Excel COM 对象 |

---

## 使用方式

本项目提供**三种使用方案**，覆盖不同用户场景：

### 方式一：下载 exe 双击运行（推荐最终用户）

访问 [Releases](https://github.com/CoreyBooks/xls2xlsx/releases) 下载 `xls2xlsx.exe`，放到目标文件夹中**双击运行**即可。无需安装 Python。

### 方式二：下载 py 脚本运行（已有 Python 环境）

如果你有 Python 环境但不方便运行 exe，可以下载 Release 中的 `xls2xlsx.py`，只需一个文件即可运行：

```bash
# 1. 安装唯一依赖
pip install pywin32>=306

# 2. 运行
python xls2xlsx.py                     # 扫描当前目录
python xls2xlsx.py D:\Data             # 扫描指定目录
python xls2xlsx.py D:\Data --delete    # 转换后删除原文件
python xls2xlsx.py D:\Data --force     # 跳过交互确认
```

| 要求 | 说明 |
|------|------|
| **Python** >= 3.9 | 运行环境 |
| **pywin32** >= 306 | `pip install pywin32` |
| **Microsoft Excel** | 2010 或更高，COM 接口必须安装 |
| **操作系统** | Windows（Excel COM 仅 Windows 可用） |

### 方式三：VBS 脚本（企业电脑受限环境）

如果你的公司电脑**无法运行外部 exe**，也**不允许安装 Python**，只需系统自带 Microsoft Excel 即可使用 VBS 版本：

1. 打开 [`VBS_vision`](./VBS_vision) 文件夹
2. 将 `xls2xlsx_en.vbs` 复制到目标文件夹
3. **双击运行**，按提示操作即可

VBS 版本零依赖、无需安装，在任何装有 Excel 的 Windows 电脑上都能直接使用。详见 [VBS_vision/README_VBS.md](./VBS_vision/README_VBS.md)。

---

### 开发者：克隆仓库

```bash
git clone https://github.com/CoreyBooks/xls2xlsx.git
cd xls2xlsx
pip install -e .
xls2xlsx    # 或 python main.py
```

---

## 打包为 exe

```bash
# 安装 PyInstaller
pip install pyinstaller

# 单文件打包
pyinstaller --onefile --name xls2xlsx --distpath ./dist main.py

# 输出: dist/xls2xlsx.exe
```

更推荐的方式：推送 Tag 后由 **GitHub Actions** 自动构建并发布到 Release。

---

## 安全与备份流程

```
[扫描预览]  统计文件数量、目录数、总大小，列出前 10 个文件
    |
[备份警告]  强制要求用户确认已完成备份（yes/no）
    |
[模式选择]  1) 保留原文件（推荐）  2) 删除原文件（危险）
    |
[最终确认]  再次确认后才启动 Excel 执行转换
    |
[执行转换]  逐文件 Workbooks.Open -> SaveAs(FileFormat=51)
    |
[结果汇总]  成功 / 跳过 / 失败分类统计
```

---

## 项目结构

```
xls2xlsx-converter-py/
├── .github/workflows/          # GitHub Actions: CI 测试 + 自动构建 Release
├── .vscode/
│   ├── launch.json             # 6 套调试配置
│   ├── settings.json           # Python 格式化、flake8 集成
│   └── extensions.json         # 推荐扩展
├── src/xls2xlsx/               # Python 版核心源码
│   ├── __init__.py
│   ├── scanner.py              # 递归扫描 .xls（严格过滤）
│   ├── converter.py            # Excel COM 生命周期管理
│   ├── cli.py                  # 彩色 CLI 交互层
│   └── main.py                 # 主流程
├── tests/                      # pytest 单元测试
│   ├── test_scanner.py
│   ├── test_converter.py
│   ├── debug_scanner.py
│   └── debug_converter.py
├── VBS_vision/                 # VBS 版本（零依赖，仅需 Excel）
│   ├── README_VBS.md           # VBS 版使用说明
│   ├── xls2xlsx.vbs            # 中文 VBS 脚本
│   └── xls2xlsx_en.vbs         # 英文 VBS 脚本（推荐分发）
├── build.spec                  # PyInstaller 打包配置
├── main.py                     # 根目录入口脚本
├── requirements.txt            # 运行时依赖（仅 pywin32）
├── requirements-dev.txt        # 开发依赖
├── setup.py                    # pip 安装配置
├── README.md
└── LICENSE
```

---

## VS Code 调试

1. 用 VS Code 打开项目根目录
2. 安装 Python 扩展（仓库已配置推荐列表）
3. 选择虚拟环境解释器
4. 按 `F5` 选择调试配置：

| 配置 | 作用 |
|------|------|
| `运行转换工具 (当前目录)` | 完整交互流程 |
| `运行转换工具 (指定测试目录)` | 扫描 `./test_data` |
| `运行转换工具 (强制删除模式)` | `--delete --force` |
| `调试扫描模块` | 独立调试 scanner.py |
| `调试转换模块` | 独立调试 converter.py |
| `运行 pytest` | 执行全部测试 |

---

## 发布到 GitHub

```bash
git add .
git commit -m "release: v1.0.0"
git tag v1.0.0
git push origin main --tags
```

GitHub Actions 自动：
- 运行 pytest 测试
- 执行 flake8 静态检查
- 用 PyInstaller 构建 `xls2xlsx.exe`
- 合并源码构建独立 `xls2xlsx.py` 脚本
- 创建 Release 并上传 exe + py 双文件

---

## 与纯 Python 方案对比

| 维度 | Python + xlrd/openpyxl | **本方案 (Python + Excel COM)** | **VBS 版** |
|------|------------------------|-----------------------------------|------------|
| 用户端依赖 | 需 Python 环境 | 需 Python + pywin32，或下载 exe | **仅需 Excel，零安装** |
| 格式保留 | 数据+基本版式（图表/宏可能丢失） | **100% 保留（Excel 自身另存为）** | **100% 保留** |
| 部署信任 | pip 安装复杂 | 单文件 exe，用户最熟悉 | **明文脚本，企业安全可控** |
| 适用场景 | 数据分析、跨平台 | 有 Python 环境 / 通用分发 | **企业受限电脑、禁止 exe 环境** |
| 跨平台 | Win/Mac/Linux | 仅 Windows | 仅 Windows |
| 工程化 | pytest, black, flake8 | **pytest, flake8, GitHub Actions 自动构建** | 手动测试为主 |

---

## 许可证

[MIT License](./LICENSE)
