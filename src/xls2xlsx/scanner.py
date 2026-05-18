#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件扫描模块：递归扫描目录中的 .xls 文件
"""
import os
from pathlib import Path
from dataclasses import dataclass, field
from typing import List


@dataclass
class ScanResult:
    """扫描结果"""
    root: Path
    files: List[Path] = field(default_factory=list)
    total_size: int = 0
    dir_count: int = 0

    @property
    def count(self) -> int:
        return len(self.files)


def scan_xls_files(root_dir: Path, recursive: bool = True) -> ScanResult:
    """
    扫描指定目录下的所有 .xls 文件（严格匹配后缀，排除 .xlsx）
    """
    if not root_dir.exists():
        raise FileNotFoundError(f"目录不存在: {root_dir}")

    pattern = "**/*.xls" if recursive else "*.xls"
    candidates = root_dir.glob(pattern)

    # 严格过滤：只保留后缀精确为 .xls 的文件
    files = sorted(
        p for p in candidates
        if p.suffix.lower() == ".xls" and p.is_file()
    )

    total_size = sum(f.stat().st_size for f in files)
    dirs = set(f.parent for f in files)

    return ScanResult(
        root=root_dir.resolve(),
        files=files,
        total_size=total_size,
        dir_count=len(dirs)
    )


def format_size(size_bytes: int) -> str:
    """字节转人类可读格式"""
    if size_bytes == 0:
        return "0 B"
    for unit in ("B", "KB", "MB", "GB"):
        if abs(size_bytes) < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} TB"
