#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scanner 模块测试
"""
import tempfile
from pathlib import Path

import pytest

from xls2xlsx.scanner import scan_xls_files, format_size


class TestScanner:
    def test_empty_directory(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            result = scan_xls_files(Path(tmpdir))
            assert result.count == 0
            assert result.total_size == 0
            assert result.dir_count == 0

    def test_find_xls_only(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            (root / "a.xls").write_text("dummy")
            (root / "b.xlsx").write_text("dummy")
            (root / "c.txt").write_text("dummy")

            result = scan_xls_files(root, recursive=False)
            assert result.count == 1
            assert result.files[0].name == "a.xls"

    def test_recursive_scan(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            (root / "root.xls").write_text("dummy")
            sub = root / "sub"
            sub.mkdir()
            (sub / "nested.xls").write_text("dummy")

            result = scan_xls_files(root, recursive=True)
            assert result.count == 2
            assert result.dir_count == 2

    def test_case_insensitive_suffix(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            (root / "upper.XLS").write_text("dummy")
            (root / "mixed.Xls").write_text("dummy")

            result = scan_xls_files(root, recursive=False)
            assert result.count == 2

    def test_format_size(self):
        assert format_size(0) == "0 B"
        assert format_size(1536) == "1.5 KB"
        assert format_size(2097152) == "2.0 MB"
