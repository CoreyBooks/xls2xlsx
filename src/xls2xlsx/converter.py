#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel COM 转换核心：基于 pywin32
调用用户本机已安装的 Microsoft Excel 执行 SaveAs(FileFormat=51)
"""
import os
import sys
import logging
from pathlib import Path
from typing import Optional, Callable

logger = logging.getLogger(__name__)

# xlOpenXMLWorkbook = 51
XL_FILE_FORMAT_XLSX = 51


class ConversionError(Exception):
    """转换异常"""
    pass


def start_excel_instance():
    """
    启动 Excel COM 实例（后台静默模式）

    Returns:
        win32com.client.Dispatch 返回的 Excel.Application 对象

    Raises:
        ConversionError: Excel 未安装或 COM 注册失败
    """
    try:
        import win32com.client as win32
    except ImportError:
        raise ConversionError(
            "缺少 pywin32 库。\n"
            "开发环境请执行: pip install pywin32\n"
            "如果已打包为 exe，此错误不应出现。"
        )

    try:
        excel = win32.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        excel.ScreenUpdating = False
        logger.info("Excel COM 实例已启动")
        return excel
    except Exception as e:
        raise ConversionError(
            f"无法启动 Microsoft Excel COM 实例。请确认已安装 Excel。\n"
            f"详细错误: {e}"
        )


def convert_single_xls(
    excel_app,
    src: Path,
    dst: Path,
    progress_callback: Optional[Callable[[int, int, str], None]] = None
):
    """
    转换单个 .xls 文件为 .xlsx

    Args:
        excel_app: 已启动的 Excel.Application COM 对象
        src: 源 .xls 路径
        dst: 目标 .xlsx 路径
        progress_callback: 进度回调 (current_sheet, total_sheets, sheet_name)

    Returns:
        dict: {"status": "Success|Skipped|Failed", "message": str}
    """
    if not src.exists():
        return {"status": "Failed", "message": f"源文件不存在: {src}"}

    if dst.exists():
        return {"status": "Skipped", "message": "目标文件已存在"}

    wb = None
    try:
        abs_src = str(src.resolve())
        abs_dst = str(dst.resolve())

        # 确保目标目录存在
        dst.parent.mkdir(parents=True, exist_ok=True)

        logger.info(f"正在打开: {abs_src}")
        wb = excel_app.Workbooks.Open(abs_src)

        # 可选：报告 Sheet 数量
        if progress_callback and wb.Sheets:
            total = wb.Sheets.Count
            for i in range(1, total + 1):
                progress_callback(i, total, wb.Sheets(i).Name)

        logger.info(f"正在保存为: {abs_dst}")
        wb.SaveAs(abs_dst, FileFormat=XL_FILE_FORMAT_XLSX)

        return {"status": "Success", "message": "转换成功"}

    except Exception as e:
        return {"status": "Failed", "message": str(e)}

    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass


def stop_excel_instance(excel_app):
    """
    安全关闭 Excel COM 实例

    先尝试 Quit，再强制垃圾回收，最大限度避免 Excel 进程残留。
    """
    if excel_app is None:
        return

    try:
        excel_app.Quit()
        logger.info("Excel.Quit() 已调用")
    except Exception as e:
        logger.warning(f"Excel.Quit() 异常: {e}")

    # 帮助 GC 回收 COM 对象
    del excel_app
    import gc
    gc.collect()
