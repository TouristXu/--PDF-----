#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
工具函数模块
包含通用的工具函数和辅助方法
"""

import os
import re
import shutil
from tkinter import messagebox


def is_valid_pdf(file_path):
    """检查文件是否为有效的PDF文件"""
    if not os.path.isfile(file_path):
        return False
    if not file_path.lower().endswith('.pdf'):
        return False
    try:
        with open(file_path, 'rb') as f:
            header = f.read(4)
            return header == b'%PDF'
    except:
        return False


def parse_page_range(range_str):
    """解析页面范围字符串"""
    pages = []
    if not range_str:
        return pages
    
    parts = range_str.split(',')
    for part in parts:
        part = part.strip()
        if '-' in part:
            try:
                start, end = part.split('-')
                start = int(start.strip())
                end = int(end.strip())
                pages.extend(range(start, end + 1))
            except ValueError:
                pass
        else:
            try:
                pages.append(int(part.strip()))
            except ValueError:
                pass
    
    return sorted(list(set(pages)))


def clean_filename(filename):
    """清理文件名中的非法字符"""
    invalid_chars = r'[\\/*?:"<>|]'
    return re.sub(invalid_chars, '_', filename)


def show_error(message):
    """显示错误消息"""
    messagebox.showerror("错误", message)


def show_info(message):
    """显示信息消息"""
    messagebox.showinfo("信息", message)


def show_warning(message):
    """显示警告消息"""
    messagebox.showwarning("警告", message)


def get_file_size(file_path):
    """获取文件大小（KB）"""
    if os.path.isfile(file_path):
        return os.path.getsize(file_path) / 1024
    return 0


def ensure_dir(path):
    """确保目录存在"""
    if path and not os.path.exists(path):
        os.makedirs(path, exist_ok=True)


def copy_file(src, dst):
    """复制文件"""
    ensure_dir(os.path.dirname(dst))
    shutil.copy2(src, dst)


def extract_name_from_filename(filename):
    """从文件名中提取姓名"""
    if not filename:
        return ""
    
    # 移除扩展名
    name = os.path.splitext(filename)[0]
    
    # 处理常见格式：姓名_编号.pdf 或 姓名-编号.pdf
    match = re.match(r'^([^_\-]+)[_\-].*$', name)
    if match:
        return match.group(1)
    
    return name


def format_page_range(pages):
    """格式化页面范围为字符串"""
    if not pages:
        return ""
    
    pages = sorted(set(pages))
    result = []
    start = pages[0]
    end = start
    
    for page in pages[1:]:
        if page == end + 1:
            end = page
        else:
            if start == end:
                result.append(str(start))
            else:
                result.append(f"{start}-{end}")
            start = page
            end = page
    
    if start == end:
        result.append(str(start))
    else:
        result.append(f"{start}-{end}")
    
    return ','.join(result)