#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
XLS -> XLSX 批量转换工具 - 根目录入口
打包为 exe 后的主入口点

使用方法:
    python main.py                    # 扫描当前目录
    python main.py D:\Data            # 扫描指定目录
    python main.py D:\Data --delete  # 转换后删除原文件
    python main.py D:\Data --force   # 跳过交互确认
"""
import sys
import argparse
from pathlib import Path

# 将 src 加入路径，支持直接运行
project_root = Path(__file__).parent.resolve()
src_path = project_root / "src"
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

from xls2xlsx.main import main


def parse_args():
    parser = argparse.ArgumentParser(
        description="XLS -> XLSX 批量转换工具 (Python + Excel COM)"
    )
    parser.add_argument(
        "path",
        nargs="?",
        default=".",
        help="扫描根目录路径 (默认: 当前目录)",
    )
    parser.add_argument(
        "--delete",
        action="store_true",
        help="转换成功后删除原 .xls 文件",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="跳过交互确认，直接执行（自动化模式）",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    main(
        root_dir=Path(args.path),
        delete_original=args.delete if args.delete else None,
        force=args.force,
    )
