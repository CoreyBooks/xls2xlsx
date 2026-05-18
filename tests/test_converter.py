#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Converter 模块测试（异常路径，无需 Excel）
"""
import tempfile
from pathlib import Path

import pytest

from xls2xlsx.converter import ConversionError


class TestConverter:
    def test_conversion_error_is_exception(self):
        with pytest.raises(ConversionError):
            raise ConversionError("test error")
