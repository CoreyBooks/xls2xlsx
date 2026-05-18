#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from setuptools import setup, find_packages

setup(
    name="xls2xlsx-converter",
    version="1.0.0",
    description="Batch XLS to XLSX Converter using Excel COM",
    author="CoreyBooks",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=["pywin32>=306"],
    python_requires=">=3.9",
    entry_points={
        "console_scripts": [
            "xls2xlsx=xls2xlsx.main:main",
        ],
    },
)
