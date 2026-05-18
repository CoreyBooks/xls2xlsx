#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
构建独立单文件 xls2xlsx.py

将 src/xls2xlsx/ 下的所有模块 + main.py 入口合并为一个独立的 .py 文件，
用户仅需 pip install pywin32 后 python xls2xlsx.py 即可运行。
"""

import re
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.resolve()
SRC_DIR = PROJECT_ROOT / "src" / "xls2xlsx"
ENTRY_FILE = PROJECT_ROOT / "main.py"
OUTPUT_DIR = PROJECT_ROOT / "dist"
OUTPUT_FILE = OUTPUT_DIR / "xls2xlsx.py"

HEADER = r'''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
XLS -> XLSX 批量转换工具 --- 独立单文件版本

使用方法:
    pip install pywin32>=306
    python xls2xlsx.py                     # 扫描当前目录
    python xls2xlsx.py D:\\Data             # 扫描指定目录
    python xls2xlsx.py D:\\Data --delete    # 转换后删除原文件
    python xls2xlsx.py D:\\Data --force     # 跳过交互确认

要求:
    - Python >= 3.9
    - Microsoft Excel 2010 或更高（Windows 必须）
    - pywin32 >= 306: pip install pywin32
"""

'''

# 模块合并顺序（注意依赖关系）
MODULE_ORDER = [
    "scanner.py",  # 无内部依赖
    "converter.py",  # 无内部依赖（除标准库）
    "cli.py",  # 依赖 scanner
    "main.py",  # 依赖 scanner, converter, cli
]


def strip_encoding_line(lines):
    """移除编码声明和 shebang 行"""
    result = []
    for line in lines:
        if line.strip().startswith("#!") and not result:
            continue
        if re.match(r"^# -\*- coding:", line.strip()):
            continue
        result.append(line)
    return result


def remove_relative_imports(lines):
    """
    将相对导入行转换为注释。
    支持单行 from .xxx import yyy 和多行 from .xxx import ( ... )
    包括函数体内的相对导入（如 cli.py 中的 from .scanner import）
    """
    result = []
    in_relative_import = False
    paren_depth = 0

    for line in lines:
        # 检测相对导入起始
        if not in_relative_import and re.match(r"^[ \t]*from \.\w+ import ", line):
            if "(" in line:
                paren_depth = line.count("(") - line.count(")")
                if paren_depth > 0:
                    in_relative_import = True
            result.append("# [MERGED] " + line)
            continue

        if in_relative_import:
            paren_depth += line.count("(") - line.count(")")
            result.append("# [MERGED] " + line)
            if paren_depth <= 0:
                in_relative_import = False
                paren_depth = 0
            continue

        result.append(line)

    return result


def strip_module_docstring(lines):
    """移除模块顶部的 docstring"""
    if not lines:
        return lines

    result = []
    i = 0
    while i < len(lines) and lines[i].strip() == "":
        result.append(lines[i])
        i += 1

    if i >= len(lines):
        return result

    line = lines[i].strip()
    if line.startswith('"""') or line.startswith("'''"):
        if line.count('"""') >= 2 or line.count("'''") >= 2:
            i += 1
        else:
            quote = line[:3]
            i += 1
            while i < len(lines):
                if quote in lines[i]:
                    i += 1
                    break
                i += 1
        while i < len(lines) and lines[i].strip() == "":
            i += 1

    result.extend(lines[i:])
    return result


def process_module(filepath):
    """读取并处理一个模块文件"""
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()

    lines = content.splitlines(keepends=True)
    lines = strip_encoding_line(lines)
    lines = strip_module_docstring(lines)
    lines = remove_relative_imports(lines)
    return "".join(lines)


def build_standalone():
    """主构建函数"""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    merged_sections = []

    # 处理 xls2xlsx 包内的各模块
    for filename in MODULE_ORDER:
        filepath = SRC_DIR / filename
        if not filepath.exists():
            print(f"[警告] 模块文件不存在: {filepath}")
            continue

        module_name = filepath.stem
        print(f"[合并] {filename} -> 模块: {module_name}")

        module_content = process_module(filepath)

        section = "\n# " + "=" * 60 + "\n"
        section += "# 模块: " + module_name + "\n"
        section += "# " + "=" * 60 + "\n\n"
        section += module_content.rstrip() + "\n"

        merged_sections.append(section)

    # 添加入口逻辑（从 root main.py 提取 argparse 部分）
    print("[合并] 入口: main.py (根目录)")

    with open(ENTRY_FILE, "r", encoding="utf-8") as f:
        entry_content = f.read()

    entry_lines = entry_content.splitlines(keepends=True)
    entry_lines = strip_encoding_line(entry_lines)
    entry_lines = strip_module_docstring(entry_lines)

    # 移除 sys.path 操作（合并后不再需要）
    entry_text = "".join(entry_lines)
    entry_text = re.sub(
        r"# 将 src 加入路径[\s\S]*?sys\.path\.insert\(0, str\(src_path\)\)\s*",
        "",
        entry_text,
    )

    # 移除 "from xls2xlsx.main import main" 导入行
    entry_text = re.sub(
        r"^from xls2xlsx\.main import main\s*$",
        "# [MERGED] main() 已在上面定义",
        entry_text,
        flags=re.MULTILINE,
    )

    merged_sections.append(
        "\n# " + "=" * 60 + "\n" "# 入口逻辑\n" "# " + "=" * 60 + "\n"
    )
    merged_sections.append(entry_text)

    # 组装最终文件
    final_content = HEADER + "".join(merged_sections)

    # 清理多余空行（连续3个以上空行压缩为2个）
    final_content = re.sub(r"\n{4,}", "\n\n\n", final_content)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(final_content)

    print(f"\n[完成] 独立脚本已生成: {OUTPUT_FILE}")
    print(f"  文件大小: {OUTPUT_FILE.stat().st_size:,} 字节")


if __name__ == "__main__":
    build_standalone()
