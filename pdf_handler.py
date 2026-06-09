#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PDF处理模块
包含PDF文件的读取、写入、合并、分割等操作
"""

import os
import PyPDF2
from utils import ensure_dir, clean_filename


class PDFHandler:
    """PDF处理类"""
    
    @staticmethod
    def get_page_count(pdf_path):
        """获取PDF页数"""
        try:
            with open(pdf_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                return len(reader.pages)
        except Exception as e:
            return 0
    
    @staticmethod
    def extract_pages(pdf_path, page_numbers, output_path):
        """提取指定页面"""
        try:
            with open(pdf_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                total_pages = len(reader.pages)
                
                writer = PyPDF2.PdfWriter()
                for page_num in page_numbers:
                    if 1 <= page_num <= total_pages:
                        writer.add_page(reader.pages[page_num - 1])
                
                ensure_dir(os.path.dirname(output_path))
                with open(output_path, 'wb') as f_out:
                    writer.write(f_out)
                return True
        except Exception as e:
            return False
    
    @staticmethod
    def merge_pdfs(pdf_paths, output_path):
        """合并多个PDF文件"""
        try:
            merger = PyPDF2.PdfMerger()
            for pdf_path in pdf_paths:
                merger.append(pdf_path)
            
            ensure_dir(os.path.dirname(output_path))
            merger.write(output_path)
            merger.close()
            return True
        except Exception as e:
            return False
    
    @staticmethod
    def append_pdfs(main_pdf_path, append_pdf_paths, output_path):
        """在主PDF文档末尾追加多个PDF文件的所有页"""
        try:
            # 验证主PDF是否有效
            if not os.path.exists(main_pdf_path):
                return False, f"主PDF文件不存在: {main_pdf_path}"
            
            merger = PyPDF2.PdfMerger()
            # 首先添加主PDF
            merger.append(main_pdf_path)
            # 然后追加所有子PDF
            for pdf_path in append_pdf_paths:
                if pdf_path:  # 确保路径不为空
                    if not os.path.exists(pdf_path):
                        return False, f"追加PDF文件不存在: {pdf_path}"
                    merger.append(pdf_path)
            
            ensure_dir(os.path.dirname(output_path))
            merger.write(output_path)
            merger.close()
            return True, "成功"
        except PyPDF2.errors.PdfReadError as e:
            return False, f"PDF文件损坏或格式错误: {str(e)}"
        except MemoryError:
            return False, "内存不足，请减少合并的文件数量或分批次处理"
        except Exception as e:
            return False, f"合并失败: {str(e)}"
    
    @staticmethod
    def append_single_page(main_pdf_path, source_pdf_path, page_number, output_path):
        """从源PDF提取指定页并追加到主PDF末尾"""
        try:
            # 验证主PDF是否有效
            if not os.path.exists(main_pdf_path):
                return False, f"主PDF文件不存在: {main_pdf_path}"
            
            # 验证源PDF是否有效
            if not os.path.exists(source_pdf_path):
                return False, f"源PDF文件不存在: {source_pdf_path}"
            
            # 创建合并器
            merger = PyPDF2.PdfMerger()
            # 添加主PDF
            merger.append(main_pdf_path)
            
            # 从源PDF提取指定页
            with open(source_pdf_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                total_pages = len(reader.pages)
                
                if page_number < 1 or page_number > total_pages:
                    return False, f"页码 {page_number} 超出范围，该PDF只有 {total_pages} 页"
                
                # 创建临时writer保存提取的页面
                temp_writer = PyPDF2.PdfWriter()
                temp_writer.add_page(reader.pages[page_number - 1])
                
                # 将提取的页面追加到合并器
                import io
                temp_buffer = io.BytesIO()
                temp_writer.write(temp_buffer)
                temp_buffer.seek(0)
                
                merger.append(temp_buffer)
            
            ensure_dir(os.path.dirname(output_path))
            merger.write(output_path)
            merger.close()
            return True, "成功"
        except PyPDF2.errors.PdfReadError as e:
            return False, f"PDF文件损坏或格式错误: {str(e)}"
        except MemoryError:
            return False, "内存不足"
        except Exception as e:
            return False, f"处理失败: {str(e)}"
    
    @staticmethod
    def split_pdf(pdf_path, split_pages, output_dir, names=None):
        """按页数分割PDF
        
        Args:
            pdf_path: 要分割的PDF文件路径
            split_pages: 每次分割的页数
            output_dir: 输出目录
            names: 可选的命名列表，用于自定义输出文件名
        """
        try:
            with open(pdf_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                total_pages = len(reader.pages)
                
                base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                start_page = 0
                part_num = 1
                
                while start_page < total_pages:
                    end_page = min(start_page + split_pages, total_pages)
                    writer = PyPDF2.PdfWriter()
                    
                    for i in range(start_page, end_page):
                        writer.add_page(reader.pages[i])
                    
                    # 确定输出文件名
                    if names and part_num <= len(names):
                        # 使用自定义名称
                        output_name = f"{names[part_num - 1]}.pdf"
                    else:
                        # 使用数字编号
                        output_name = f"{base_name}_part{part_num}.pdf"
                    
                    output_path = os.path.join(output_dir, output_name)
                    with open(output_path, 'wb') as f_out:
                        writer.write(f_out)
                    
                    start_page = end_page
                    part_num += 1
                
                return True, f"成功分割为 {part_num - 1} 个文件"
        except Exception as e:
            return False, f"分割失败: {str(e)}"
    
    @staticmethod
    def swap_pages(pdf_path, page_a, page_b, output_path):
        """交换PDF中的两个页面"""
        try:
            with open(pdf_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                total_pages = len(reader.pages)
                
                if page_a < 1 or page_a > total_pages or page_b < 1 or page_b > total_pages:
                    return False
                
                writer = PyPDF2.PdfWriter()
                for i in range(total_pages):
                    if i == page_a - 1:
                        writer.add_page(reader.pages[page_b - 1])
                    elif i == page_b - 1:
                        writer.add_page(reader.pages[page_a - 1])
                    else:
                        writer.add_page(reader.pages[i])
                
                ensure_dir(os.path.dirname(output_path))
                with open(output_path, 'wb') as f_out:
                    writer.write(f_out)
                return True
        except Exception as e:
            return False
    
    @staticmethod
    def reorder_pages(pdf_path, page_order, output_path):
        """按照指定顺序重排PDF页面"""
        try:
            with open(pdf_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                total_pages = len(reader.pages)
                
                writer = PyPDF2.PdfWriter()
                for page_num in page_order:
                    if 1 <= page_num <= total_pages:
                        writer.add_page(reader.pages[page_num - 1])
                
                ensure_dir(os.path.dirname(output_path))
                with open(output_path, 'wb') as f_out:
                    writer.write(f_out)
                return True
        except Exception as e:
            return False
    
    @staticmethod
    def insert_pages(source_pdf, insert_pdf, position, output_path):
        """在指定位置插入页面"""
        try:
            merger = PyPDF2.PdfMerger()
            merger.append(source_pdf)
            merger.merge(position - 1, insert_pdf)
            
            ensure_dir(os.path.dirname(output_path))
            merger.write(output_path)
            merger.close()
            return True
        except Exception as e:
            return False
    
    @staticmethod
    def reverse_pages(pdf_path, output_path):
        """反转PDF页面顺序"""
        try:
            with open(pdf_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                writer = PyPDF2.PdfWriter()
                
                for page in reversed(reader.pages):
                    writer.add_page(page)
                
                ensure_dir(os.path.dirname(output_path))
                with open(output_path, 'wb') as f_out:
                    writer.write(f_out)
                return True
        except Exception as e:
            return False
    
    @staticmethod
    def rotate_pages(pdf_path, rotation_scheme, output_path):
        """按指定方案旋转页面"""
        try:
            with open(pdf_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                writer = PyPDF2.PdfWriter()
                
                for i, page in enumerate(reader.pages):
                    if i + 1 in rotation_scheme:
                        page.rotate(rotation_scheme[i + 1])
                    writer.add_page(page)
                
                ensure_dir(os.path.dirname(output_path))
                with open(output_path, 'wb') as f_out:
                    writer.write(f_out)
                return True
        except Exception as e:
            return False