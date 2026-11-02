#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""通用工具函数。"""


def format_time(seconds: float) -> str:
    """格式化时间显示。

    Args:
        seconds: 秒数

    Returns:
        格式化的时间字符串，如 "5.23秒"、"2分30.50秒"、"1时30分45.00秒"
    """
    if seconds < 60:
        return f"{seconds:.2f}秒"
    if seconds < 3600:
        minutes = int(seconds // 60)
        secs = seconds % 60
        return f"{minutes}分{secs:.2f}秒"
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = seconds % 60
    return f"{hours}时{minutes}分{secs:.2f}秒"
