#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
命令行交互模块：彩色输出、扫描报告、警告、确认流程
"""

import sys
from pathlib import Path


# ANSI 颜色码（Windows 10+ 和 VS Code 终端均支持）
class Color:
    CYAN = "\033[96m"
    YELLOW = "\033[93m"
    RED = "\033[91m"
    GREEN = "\033[92m"
    WHITE = "\033[97m"
    MAGENTA = "\033[95m"
    DARK_GRAY = "\033[90m"
    RESET = "\033[0m"


def _c(text: str, color: str) -> str:
    return f"{color}{text}{Color.RESET}"


def show_banner():
    line = "=" * 64
    print(f"\n{_c(line, Color.CYAN)}")
    print(_c("        XLS -> XLSX 批量转换工具 (Python + Excel COM)", Color.CYAN))
    print(_c(line, Color.CYAN))


def show_features():
    print(_c("【技术特点】", Color.WHITE))
    print(_c("  . 基于 Microsoft Excel COM 接口，保留 100% 原始格式", Color.WHITE))
    print(_c("  . 打包为单文件 exe，无需安装 Python", Color.WHITE))
    print(_c("  . 递归扫描目录，批量处理，自动跳过已存在文件", Color.WHITE))
    print(_c("  . 完善的备份警告与确认机制，防止误操作", Color.WHITE))
    print(_c("-" * 64, Color.CYAN))


def show_scan_report(scan_result):
    from .scanner import format_size

    print(f"\n{_c('[扫描结果]', Color.MAGENTA)}")
    print(_c(f"  待转换文件 : {scan_result.count} 个", Color.WHITE))
    print(_c(f"  涉及目录   : {scan_result.dir_count} 个", Color.WHITE))
    print(_c(f"  总大小     : {format_size(scan_result.total_size)}", Color.WHITE))

    max_show = min(10, scan_result.count)
    print(f"\n{_c(f'[文件清单] 前 {max_show} 个:', Color.MAGENTA)}")
    for i, f in enumerate(scan_result.files[:max_show], 1):
        try:
            rel = f.relative_to(scan_result.root)
        except ValueError:
            rel = f
        print(_c(f"  {i:2d}. {rel}", Color.WHITE))
    if scan_result.count > 10:
        print(_c(f"      ... 还有 {scan_result.count - 10} 个文件", Color.DARK_GRAY))


def show_warning():
    line = "!" * 64
    print(f"\n{_c(line, Color.YELLOW)}")
    print(_c("  !   重 要 警 告", Color.YELLOW))
    print(_c(line, Color.YELLOW))
    print(_c("  1. 本工具通过 Excel COM 接口执行另存为操作，", Color.YELLOW))
    print(_c("     转换质量与手动在 Excel 中'另存为'完全一致。", Color.YELLOW))
    print(_c("  2. 转换过程会生成新的 .xlsx 文件，不会直接覆盖原 .xls。", Color.YELLOW))
    print(_c("  3. 但若选择'删除原文件'模式，原始 .xls 将被永久删除！", Color.YELLOW))
    print(_c("  4. 强烈建议在运行前对原始文件进行完整备份！", Color.YELLOW))
    print(_c(line, Color.YELLOW))


def confirm_backup() -> bool:
    while True:
        ans = (
            input(
                "\n是否已完成原始文件备份？\n"
                "  [1] 已备份\n"
                "  [Q] 未备份，退出程序\n"
                "请输入: "
            )
            .strip()
            .lower()
        )
        if ans == "1":
            return True
        elif ans == "q":
            print(_c("\n[退出] 请先备份原始文件后再运行。", Color.YELLOW))
            sys.exit(0)
        print(_c("  >> 输入无效，请输入 1 或 Q", Color.RED))


def choose_mode() -> bool:
    print(f"\n{_c('【操作选项】', Color.MAGENTA)}")
    print(_c("  1. 转换后保留原 .xls 文件（安全，推荐）", Color.WHITE))
    print(_c("  2. 转换后删除原 .xls 文件（危险，仅确认备份后使用）", Color.RED))

    while True:
        choice = input("\n请输入选项 (1 或 2): ").strip()
        if choice == "1":
            return False
        elif choice == "2":
            return True
        print(_c("  >> 输入无效，请重新输入。", Color.RED))


def confirm_execution(count: int, delete_original: bool) -> bool:
    mode_str = "【删除原文件】" if delete_original else "【保留原文件】"
    print(
        f"\n{_c(f'[执行确认] {mode_str} 即将处理 {count} 个 .xls 文件', Color.MAGENTA)}"
    )

    while True:
        ans = (
            input(
                "是否开始转换？\n"
                "  [1] 确认开始\n"
                "  [2] 返回上一级\n"
                "  [Q] 退出程序\n"
                "请输入: "
            )
            .strip()
            .lower()
        )
        if ans == "1":
            return True
        elif ans == "2":
            return False
        elif ans == "q":
            print(_c("\n[退出] 操作已取消。", Color.YELLOW))
            sys.exit(0)
        print(_c("  >> 输入无效，请输入 1、2 或 Q", Color.RED))


def show_progress(current: int, total: int, filename: str):
    pct = current / total * 100 if total > 0 else 0
    print(f"\r  [{current:3d}/{total}] {pct:5.1f}% - {filename}", end="", flush=True)


def show_summary(success: int, skipped: int, failed: list):
    print(f"\n\n{_c('=' * 64, Color.CYAN)}")
    print(_c("                     处 理 结 果 汇 总", Color.CYAN))
    print(_c("=" * 64, Color.CYAN))
    print(_c(f"  . 成功转换 : {success} 个", Color.GREEN))
    print(_c(f"  . 跳过(已存在): {skipped} 个", Color.MAGENTA))
    print(_c(f"  . 转换失败 : {len(failed)} 个", Color.RED))

    if failed:
        print(_c("-" * 64, Color.CYAN))
        print(_c("【失败明细】", Color.RED))
        for item in failed:
            print(_c(f"  . {item['source']}", Color.RED))
            print(_c(f"    原因: {item['message']}", Color.DARK_GRAY))
    print(_c("=" * 64, Color.CYAN))


def pause_exit(code: int = 0):
    try:
        input("\n按回车键退出...")
    except (KeyboardInterrupt, EOFError):
        pass
    sys.exit(code)
