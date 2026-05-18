#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主流程：扫描 -> 确认 -> 转换 -> 汇总
"""

import sys
import logging
from pathlib import Path

from .scanner import scan_xls_files
from .converter import (
    start_excel_instance,
    convert_single_xls,
    stop_excel_instance,
    ConversionError,
)
from .cli import (
    show_banner,
    show_features,
    show_scan_report,
    show_warning,
    confirm_backup,
    choose_mode,
    confirm_execution,
    show_progress,
    show_summary,
    pause_exit,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


def main(root_dir: Path = None, delete_original: bool = None, force: bool = False):
    """
    主入口

    Args:
        root_dir: 扫描根目录，默认当前工作目录
        delete_original: 是否删除原文件，None 表示由交互选择
        force: 是否跳过所有交互确认（自动化模式）
    """
    if root_dir is None:
        root_dir = Path.cwd()

    show_banner()
    show_features()

    print(f"[扫描目录] {root_dir.resolve()}\n")
    print("正在扫描 .xls 文件...")

    try:
        scan = scan_xls_files(root_dir)
    except Exception as e:
        print(f"\n[错误] 扫描失败: {e}")
        pause_exit(1)
        return

    if scan.count == 0:
        print("\n[结果] 未发现任何 .xls 文件，程序退出。")
        pause_exit(0)
        return

    # 扫描报告
    show_scan_report(scan)

    # 交互确认（非 force 模式）
    if not force:
        show_warning()
        confirm_backup()  # 按 Q 会直接 sys.exit()

        if delete_original is None:
            delete_original = choose_mode()

        while True:
            if confirm_execution(scan.count, delete_original):
                break  # 用户确认开始转换
            # 用户选择 [2] 返回上一级（模式选择）
            delete_original = choose_mode()

    # 启动 Excel
    print("\n正在启动 Microsoft Excel...")
    excel = None
    try:
        excel = start_excel_instance()
    except ConversionError as e:
        print(f"\n[错误] {e}")
        pause_exit(1)
        return

    # 执行转换
    print("-" * 64)
    print("开始转换...\n")

    success = 0
    skipped = 0
    failed = []

    for i, src_file in enumerate(scan.files, 1):
        dst_path = src_file.with_suffix(".xlsx")

        show_progress(i, scan.count, src_file.name)

        result = convert_single_xls(excel, src_file, dst_path)

        if result["status"] == "Success":
            success += 1
            if delete_original:
                try:
                    src_file.unlink()
                except Exception as e:
                    logger.warning(f"无法删除原文件: {src_file} - {e}")
        elif result["status"] == "Skipped":
            skipped += 1
        else:
            failed.append({"source": str(src_file), "message": result["message"]})

    print()  # 结束进度行

    # 汇总
    show_summary(success, skipped, failed)

    # 清理
    if excel:
        stop_excel_instance(excel)

    pause_exit(1 if failed else 0)
