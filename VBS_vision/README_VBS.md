# XLS -> XLSX 批量转换工具（VBScript 版）

`xls2xlsx` 的独立 VBScript 实现。不需要安装 Python，只需系统上装有 Microsoft Excel 即可运行。

> **范围：** 本 README 针对 VBScript 版本（`xls2xlsx_en.vbs`）。Python 版本请查看主项目 [`README.md`](./README.md)。

---

## 目录

- [功能特性](#功能特性)
- [系统要求](#系统要求)
- [快速开始](#快速开始)
- [用户交互流程](#用户交互流程)
- [打包与分发](#打包与分发)
- [技术要点](#技术要点)
- [与 Python 版本的对比](#与-python-版本的对比)

---

## 功能特性

| 特性 | 说明 |
|------|------|
| **100% 格式保留** | 调用 Excel 原生 COM `SaveAs(FileFormat=51)`，效果和手动在 Excel 中"另存为"完全一致。 |
| **递归扫描** | 自动扫描脚本所在文件夹及其所有子文件夹中的 `.xls` 文件。 |
| **严格过滤** | 只处理后缀精确为 `.xls` 的文件，`.xlsx` 文件自动忽略。 |
| **默认安全** | 如果同名 `.xlsx` 已存在，自动跳过，不会覆盖。 |
| **多层确认机制** | 备份确认 -> 模式选择 -> 执行确认，防止误操作导致数据丢失。 |
| **可选清理** | 转换成功后可选择"保留原文件"（安全）或"删除原文件"（危险）。 |
| **详细汇总** | 转换结束后统计成功数、跳过数，并列出完整失败日志及原因。 |

---

## 系统要求

| 项目 | 最低要求 |
|------|---------|
| 操作系统 | Windows 7 或更高版本 |
| Microsoft Excel | 2010 或更高版本（Excel COM 必须已注册） |
| 权限 | 对目标文件夹的读写权限 |

不需要安装 Python、.NET 运行时或任何第三方库。

---

## 快速开始

### 方式一：直接使用（最简单）

1. 将 `xls2xlsx_en.vbs` 复制到存放 `.xls` 文件的文件夹中。
2. 双击运行脚本。
3. 脚本会自动检测当前运行环境是 `wscript`（GUI 模式）还是 `cscript`（控制台模式）。如果检测到 GUI 模式，会自动在后台用控制台模式重新启动自身，以便正常交互。
4. 根据屏幕提示输入：
   - `1` = 确认 / 是
   - `2` = 返回上一级
   - `Q` = 退出

### 方式二：命令行运行

```cmd
cscript //nologo xls2xlsx_en.vbs
```

> **提示：** 在某些 Windows 配置下，你可以将文件夹拖拽到脚本图标上运行，但推荐的做法还是直接把脚本放进目标目录里双击运行。

---

## 用户交互流程

```
[扫描预览]  统计文件数量、目录数、总大小，列出前 10 个文件
       |
[备份确认]  强制要求用户确认已完成备份（1 = 已备份，Q = 退出）
       |
[模式选择]  1) 保留原文件（推荐）  2) 删除原文件（危险）
       |
[最终确认]  1) 开始转换  2) 返回上一级  Q) 退出
       |
[执行转换]  后台静默启动 Excel，逐文件处理
       |
[结果汇总]  成功 / 跳过 / 失败 分类统计
```

---

## 打包与分发

`.vbs` 文件本身在已安装 Excel 的 Windows 电脑上已经是可执行文件。但如果你想把它打包成单个 `.exe`（例如为了隐藏源代码，或让它看起来更像一个"传统"应用程序），以下是三种常用方案。

### 方案一：IExpress（Windows 自带工具）-- 推荐

`IExpress.exe` 随所有 Windows 系统附带。它可以创建静默运行的自解压程序包。

**第 1 步：** 在与脚本相同的文件夹中创建一个启动批处理文件 `run.bat`：

```batch
@echo off
cscript //nologo "%~dp0xls2xlsx_en.vbs"
```

**第 2 步：** 启动 IExpress

```cmd
iexpress
```

**第 3 步：** 按向导操作
- 选择 **"Create new Self Extraction Directive file"**（创建新的自解压指令文件）
- 选择 **"Extract files and run an installation command"**（解压文件并运行安装命令）
- 将 `xls2xlsx_en.vbs` 和 `run.bat` 都添加为打包文件
- **安装程序**设置为：`run.bat`
- 选择 **"Hide Program"**（隐藏程序）和 **"No prompt"**（不提示），实现静默启动
- 保存输出 `.exe`（例如 `xls2xlsx_en.exe`）

**优点：** 无需下载额外工具；Windows 原生信任。  
**缺点：** `.vbs` 源码仍可从 `.exe` 中提取出来（不是真正的编译）。

---

### 方案二：WinRAR 自解压（需已安装 WinRAR）

1. 在 WinRAR 中选中 `xls2xlsx_en.vbs`（以及 `run.bat`，如果需要）。
2. 点击 **工具 -> 转换档案为自解压格式**。
3. 在 **高级自解压选项 -> 设置** 中，将 **解压后运行** 设为 `run.bat`。
4. 在 **模式** 中勾选 **解压到临时文件夹** 和 **全部隐藏**。
5. 生成 `.exe`。

**优点：** 如果你已经在用 WinRAR，操作非常方便。  
**缺点：** WinRAR 是付费软件（虽然试用版也支持此功能）。

---

### 方案三：第三方 "脚本转 Exe" 工具

**Bat To Exe Converter**、**VbsEdit** 或 **Advanced BAT to EXE Converter** 等工具可以将 `.vbs` 打包成独立的 `.exe`。

**一般流程：**
1. 将 `xls2xlsx_en.vbs` 导入工具。
2. 选择目标架构（x86 / x64）。
3. 设置输出文件名（例如 `xls2xlsx_en.exe`）。
4. （可选）嵌入图标、要求管理员权限等。
5. 生成。

**优点：** 真正的单文件 `.exe`；可嵌入图标和元数据。  
**缺点：** 需要下载第三方软件；部分杀毒软件可能对生成的 exe 报误报。

---

## 技术要点

### 1. CScript 与 WScript 的区别

VBScript 文件（`.vbs`）可以由两个宿主程序执行：

| 宿主 | 行为 |
|------|------|
| `wscript.exe` | 默认双击打开的宿主。`WScript.Echo` 会以弹出消息框显示。**无法正常使用 `StdIn` 交互输入。** |
| `cscript.exe` | 控制台宿主。`WScript.Echo` 输出到终端；`StdIn.ReadLine()` 可正常读取键盘输入。 |

**本脚本的解决方案：**

```vbscript
If InStr(LCase(WScript.FullName), "cscript.exe") = 0 Then
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "cscript.exe //nologo """ & WScript.ScriptFullName & """"
    Set objShell = Nothing
    WScript.Quit
End If
```

当用户双击文件（由 `wscript` 启动）时，脚本检测到这一点，立即用 `cscript` 重新启动一个控制台进程，然后自身退出。用户会看到一个命令行窗口弹出，可以正常输入交互内容。

### 2. 编码陷阱

Windows 脚本宿主（cscript/wscript）引擎默认使用**系统默认 ANSI 代码页**解析脚本文件（例如中文 Windows 使用 `CP936`/GBK，西欧 Windows 使用 `CP1252`），**除非**文件开头带有 UTF-8 BOM，或使用了 XML 编码声明。

**实际规则：**
- 如果脚本包含**非 ASCII 字符**（中文、日文、Emoji 等），必须按系统默认 ANSI 编码保存（如中文 Windows 用 GB2312）。否则 cscript 会把多字节字符错误拆分为多个独立符号，抛出"无效字符"或"未结束的字符串常量"错误。
- 如果脚本是**纯 ASCII**（英文、数字、基本标点），UTF-8 无 BOM 在任何系统上都是安全的。

**这就是为什么英文版（`xls2xlsx_en.vbs`）是推荐的分发文件** --- 它彻底避开了所有跨区域编码问题。

### 3. Excel COM 对象生命周期管理

防止 Excel 进程残留的关键设置：

```vbscript
excel.DisplayAlerts = False
excel.Visible       = False
excel.ScreenUpdating = False
```

转换结束后：

```vbscript
excel.Quit
Set excel = Nothing
```

此外，每个文件的转换操作都包裹在 `On Error Resume Next ... On Error GoTo 0` 块中。如果某个文件损坏或被密码保护导致打开失败，错误会被捕获、记录到失败列表，然后循环继续处理下一个文件，而不会导致整批任务崩溃。

### 4. 不用 `System.Collections.ArrayList` 实现动态数组

VBScript 没有原生动态数组类型。本脚本使用 `Scripting.Dictionary` 配合整数键来实现轻量级的动态列表：

```vbscript
Dim list
Set list = CreateObject("Scripting.Dictionary")
list.Add list.Count, "item1"   ' 索引 0
list.Add list.Count, "item2"   ' 索引 1
```

这避免了对 .NET COM 对象（如 `System.Collections.ArrayList`）的依赖，确保在精简版或较老的 Windows 系统上也能正常运行。

### 5. 文件格式常量

```vbscript
Const XL_FILE_FORMAT_XLSX = 51
```

`51` 是 Excel 对象模型中 `xlOpenXMLWorkbook` 的枚举值。将它传给 `Workbook.SaveAs` 可以强制输出为现代 `.xlsx` 格式，无论输入文件是什么格式。

### 6. 严格的后缀检查

```vbscript
If LCase(fso.GetExtensionName(file.Name)) = "xls" Then
```

这确保 `file.xls` 会被匹配，而 `file.xlsx` **不会**。如果简单地用 `Like "*.xls"` 模式匹配，就会错误地把 `.xlsx` 文件也包含进来。

---

## 与 Python 版本的对比

| 维度 | Python + pywin32 | **VBScript** |
|------|------------------|--------------|
| 运行时依赖 | 需要安装 Python + pywin32 | **无需额外安装（Windows 内置）** |
| 单文件分发 | 通过 PyInstaller 打包为 `.exe` | **`.vbs` 本身就是可执行文件** |
| 源码可见性 | 编译为 exe 后不可见 | 明文可见 |
| 跨平台 | 仅限 Windows（COM 限制） | 仅限 Windows（COM 限制） |
| 格式保留 | 完全一致（均使用 Excel COM `SaveAs`） | 完全一致 |
| 错误处理 | 丰富（完整堆栈跟踪、日志系统） | 基础（Err.Number / Err.Description） |
| 工程化能力 | 支持 pytest、CI/CD、静态检查 | 手动测试为主 |

**适合使用 VBScript 版本的场景：**
- 目标用户**没有** Python 环境，且不方便分发体积较大的 `.exe`。
- 需要一个**零依赖**方案，在任何装有 Excel 的企业 Windows 电脑上都能即开即用。
- 希望源码能被非开发人员（IT 运维、财务人员）直接阅读甚至修改。

**适合使用 Python 版本的场景：**
- 需要自动化测试、CI/CD 集成或高级错误处理。
- 希望分发一个编译后的单文件 `.exe`，隐藏具体实现。
- 计划扩展更多功能（例如写入数据库日志、上传云端、导出 PDF 等）。

---

## 许可证

与主项目一致：[MIT License](./LICENSE)
