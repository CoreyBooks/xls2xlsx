#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
扫描模块独立调试脚本
无需 Excel，可单独运行测试扫描逻辑
"""
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent.resolve()
src_path = project_root / "src"
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

from xls2xlsx.scanner import scan_xls_files

# 创建临时测试数据
test_root = project_root / "test_data"
if not test_root.exists():
    test_root.mkdir()
    (test_root / "a.xls").write_text("dummy")
    (test_root / "b.xlsx").write_text("dummy")
    (test_root / "sub").mkdir()
    (test_root / "sub" / "nested.xls").write_text("dummy")

result = scan_xls_files(test_root)
print(f"扫描完成: {result.count} 个文件")
for f in result.files:
    print(f"  - {f.name} ({f.stat().st_size} bytes)")
