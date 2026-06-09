#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Office转PDF模块
包含Word和Excel转PDF的功能
"""

import os
import re
from utils import ensure_dir, clean_filename


class OfficeConverter:
    """Office文档转PDF转换器"""
    
    @staticmethod
    def word_to_pdf(word_path, output_path, page_range=None):
        """将Word文档转换为PDF"""
        try:
            import win32com.client as win32
            from win32com.client import constants
            
            word = win32.gencache.EnsureDispatch('Word.Application')
            word.Visible = False
            
            try:
                doc = word.Documents.Open(word_path, ReadOnly=True)
                
                # 设置导出参数
                export_params = {
                    'ExportFormat': 17,  # wdExportFormatPDF
                    'OpenAfterExport': False,
                    'OptimizeFor': 0,  # wdExportOptimizeForPrint
                    'Range': 0,  # wdExportAllDocument
                    'From': 1,
                    'To': 1
                }
                
                # 如果指定了页面范围
                if page_range and page_range.strip():
                    pages = OfficeConverter._parse_page_range(page_range)
                    if pages:
                        export_params['Range'] = 3  # wdExportFromTo
                        export_params['From'] = min(pages)
                        export_params['To'] = max(pages)
                
                ensure_dir(os.path.dirname(output_path))
                doc.ExportAsFixedFormat(output_path, **export_params)
                
                doc.Close(SaveChanges=False)
                return True
            finally:
                word.Quit()
        except Exception as e:
            return False
    
    @staticmethod
    def excel_to_pdf(excel_path, output_path, sheets=None, orientation='landscape', scale_mode='fit_width', scale=70,
                     print_active_sheet=False, paper_size='A4', first_page=None, last_page=None):
        """将Excel文档转换为PDF
        
        Args:
            excel_path: Excel文件路径
            output_path: 输出PDF路径
            sheets: 指定要导出的Sheet列表
            orientation: 页面方向 'landscape' 横向 / 'portrait' 纵向
            scale_mode: 缩放模式 'fit_width'适应宽度 / 'normal'正常 / 'custom'自定义
            scale: 自定义缩放比例（当scale_mode='custom'时使用）
            print_active_sheet: 是否只打印活动工作表
            paper_size: 纸张大小 'A4', 'A3', 'Letter' 等
            first_page: 起始页码（None表示从第一页开始）
            last_page: 结束页码（None表示到最后一页）
        """
        try:
            import win32com.client as win32
            from win32com.client import constants
            
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            
            try:
                workbook = excel.Workbooks.Open(excel_path, ReadOnly=True)
                
                # 纸张大小映射
                paper_size_map = {
                    'A4': 9,      # xlPaperA4
                    'A3': 8,      # xlPaperA3
                    'Letter': 1,  # xlPaperLetter
                    'Legal': 5,   # xlPaperLegal
                    'B4': 12,     # xlPaperB4
                    'B5': 13,     # xlPaperB5
                }
                
                # 如果指定了Sheet，隐藏其他Sheet
                if sheets and isinstance(sheets, list) and len(sheets) > 0:
                    # 先保存所有Sheet的可见状态
                    original_visibility = {}
                    for sheet in workbook.Sheets:
                        original_visibility[sheet.Name] = sheet.Visible
                        sheet.Visible = False
                    
                    # 只显示指定的Sheet
                    for sheet_name in sheets:
                        try:
                            workbook.Sheets(sheet_name).Visible = True
                        except:
                            pass
                
                # 设置所有可见Sheet的打印参数
                for sheet in workbook.Sheets:
                    if sheet.Visible:
                        # 设置页面方向
                        sheet.PageSetup.Orientation = 2 if orientation == 'landscape' else 1  # xlLandscape=2, xlPortrait=1
                        # 设置纸张大小
                        sheet.PageSetup.PaperSize = paper_size_map.get(paper_size, 9)  # 默认A4
                        
                        # 设置缩放模式 - 必须先设置Zoom，再设置FitToPages属性
                        try:
                            if scale_mode == 'fit_width':
                                # 适应宽度（列不拆分）
                                sheet.PageSetup.Zoom = False
                                sheet.PageSetup.FitToPagesWide = 1
                                # 不设置FitToPagesTall，让Excel自动处理
                            elif scale_mode == 'fit_all':
                                # 适应页面（所有内容在一页）
                                sheet.PageSetup.Zoom = False
                                sheet.PageSetup.FitToPagesWide = 1
                                sheet.PageSetup.FitToPagesTall = 1
                            elif scale_mode == 'normal':
                                # 正常大小
                                sheet.PageSetup.Zoom = True
                                sheet.PageSetup.ZoomPercent = 100
                            elif scale_mode == 'custom':
                                # 自定义缩放比例
                                sheet.PageSetup.Zoom = True
                                sheet.PageSetup.ZoomPercent = int(scale)
                        except Exception as e:
                            # PageSetup设置失败时忽略，使用默认设置
                            pass
                
                ensure_dir(os.path.dirname(output_path))
                
                # 使用位置参数调用ExportAsFixedFormat
                # ExportAsFixedFormat(Type, Filename, Quality, IncludeDocProperties, 
                #                     IgnorePrintAreas, From, To, OpenAfterExport, FixedFormatExtClassPtr)
                export_type = 0  # xlTypePDF
                quality = 0      # xlQualityStandard
                include_doc_props = True
                ignore_print_areas = False
                open_after_export = False
                
                # 如果只打印活动工作表
                if print_active_sheet:
                    # 只导出活动工作表
                    active_sheet = workbook.ActiveSheet
                    if first_page is not None and last_page is not None:
                        active_sheet.ExportAsFixedFormat(
                            export_type, output_path, quality, include_doc_props,
                            ignore_print_areas, first_page, last_page, open_after_export
                        )
                    elif first_page is not None:
                        active_sheet.ExportAsFixedFormat(
                            export_type, output_path, quality, include_doc_props,
                            ignore_print_areas, first_page, 9999, open_after_export
                        )
                    elif last_page is not None:
                        active_sheet.ExportAsFixedFormat(
                            export_type, output_path, quality, include_doc_props,
                            ignore_print_areas, 1, last_page, open_after_export
                        )
                    else:
                        active_sheet.ExportAsFixedFormat(
                            export_type, output_path, quality, include_doc_props,
                            ignore_print_areas
                        )
                else:
                    # 导出整个工作簿
                    if first_page is not None and last_page is not None:
                        workbook.ExportAsFixedFormat(
                            export_type, output_path, quality, include_doc_props,
                            ignore_print_areas, first_page, last_page, open_after_export
                        )
                    elif first_page is not None:
                        workbook.ExportAsFixedFormat(
                            export_type, output_path, quality, include_doc_props,
                            ignore_print_areas, first_page, 9999, open_after_export
                        )
                    elif last_page is not None:
                        workbook.ExportAsFixedFormat(
                            export_type, output_path, quality, include_doc_props,
                            ignore_print_areas, 1, last_page, open_after_export
                        )
                    else:
                        workbook.ExportAsFixedFormat(
                            export_type, output_path, quality, include_doc_props,
                            ignore_print_areas
                        )
                
                # 恢复Sheet可见性
                if sheets:
                    for sheet_name, visible in original_visibility.items():
                        try:
                            workbook.Sheets(sheet_name).Visible = visible
                        except:
                            pass
                
                workbook.Close(SaveChanges=False)
                return True
            finally:
                excel.Quit()
        except Exception as e:
            return False
    
    @staticmethod
    def _parse_page_range(range_str):
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
    
    @staticmethod
    def batch_word_to_pdf(word_files, output_dir, page_range=None, progress_callback=None):
        """批量将Word文档转换为PDF"""
        results = {
            'success': 0,
            'failed': 0,
            'fail_messages': []
        }
        
        total_files = len(word_files)
        for i, word_path in enumerate(word_files):
            try:
                base_name = os.path.splitext(os.path.basename(word_path))[0]
                base_name = clean_filename(base_name)
                
                if page_range and page_range.strip():
                    output_filename = f"{base_name}_pages_{page_range.replace(',', '_').replace('-', '_')}.pdf"
                else:
                    output_filename = f"{base_name}.pdf"
                
                output_path = os.path.join(output_dir, output_filename)
                
                if OfficeConverter.word_to_pdf(word_path, output_path, page_range):
                    results['success'] += 1
                else:
                    results['failed'] += 1
                    results['fail_messages'].append(f"{os.path.basename(word_path)}: 转换失败")
            except Exception as e:
                results['failed'] += 1
                results['fail_messages'].append(f"{os.path.basename(word_path)}: {str(e)}")
            
            if progress_callback:
                progress_callback((i + 1) / total_files * 100)
        
        return results
    
    @staticmethod
    def batch_excel_to_pdf(excel_files, output_dir, sheets=None, progress_callback=None, 
                           orientation='landscape', scale_mode='fit_width', scale=70,
                           print_active_sheet=False, paper_size='A4', first_page=None, last_page=None):
        """批量将Excel文档转换为PDF"""
        results = {
            'success': 0,
            'failed': 0,
            'fail_messages': []
        }
        
        total_files = len(excel_files)
        for i, excel_path in enumerate(excel_files):
            try:
                base_name = os.path.splitext(os.path.basename(excel_path))[0]
                base_name = clean_filename(base_name)
                output_filename = f"{base_name}.pdf"
                output_path = os.path.join(output_dir, output_filename)
                
                if OfficeConverter.excel_to_pdf(excel_path, output_path, sheets, 
                                                orientation, scale_mode, scale,
                                                print_active_sheet, paper_size,
                                                first_page, last_page):
                    results['success'] += 1
                else:
                    results['failed'] += 1
                    results['fail_messages'].append(f"{os.path.basename(excel_path)}: 转换失败")
            except Exception as e:
                results['failed'] += 1
                results['fail_messages'].append(f"{os.path.basename(excel_path)}: {str(e)}")
            
            if progress_callback:
                progress_callback((i + 1) / total_files * 100)
        
        return results