#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
转换模块独立调试脚本
需要真实 .xls 文件和已安装的 Excel
"""
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent.resolve()
src_path = project_root / "src"
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

from xls2xlsx.converter import start_excel_instance, convert_single_xls, stop_excel_instance

test_root = project_root / "test_data"
sample = list(test_root.glob("*.xls"))

if not sample:
    print("请在 test_data/ 中放入真实 .xls 文件")
    sys.exit(1)

src = sample[0]
dst = test_root / "output" / "debug.xlsx"

print(f"样本: {src}")
print("启动 Excel...")
excel = start_excel_instance()

print("执行转换...")
result = convert_single_xls(excel, src, dst)
print(f"结果: {result}")

print("关闭 Excel...")
stop_excel_instance(excel)
print("完成")
