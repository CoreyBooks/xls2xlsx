#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
包入口: python -m xls2xlsx
"""
import sys
from pathlib import Path
from .main import main

if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    main(target)
