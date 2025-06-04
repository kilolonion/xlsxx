# -*- coding: utf-8 -*-
"""
文件转换器模块
实现XLS/XLSX文件到各种格式的转换
"""

import os
import pandas as pd
from typing import Optional, Dict, Any, List
from pathlib import Path
import streamlit as st
from io import BytesIO

# PDF生成相关导入
from reportlab.lib.pagesizes import letter, A4, A3
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import inch

from config.settings import PDF_PAGE_SIZES, CSV_SEPARATORS


class ConversionError(Exception):
    """转换错误异常类"""
    pass


class FileConverter:
    """文件转换器类"""
    
    def __init__(self):
        """初始化转换器"""
        self.supported_input_formats = ['xls', 'xlsx']
        self.supported_output_formats = ['csv', 'pdf', 'xlsx']
    
    def convert_to_csv(self, input_path: str, output_path: str, 
                      encoding: str = 'utf-8', separator: str = ',',
                      sheet_name: Optional[str] = None) -> bool:
        """
        转换Excel文件为CSV格式
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            encoding: 输出编码
            separator: 分隔符
            sheet_name: 工作表名称，None表示第一个工作表
            
        Returns:
            转换是否成功
        """
        try:
            # 读取Excel文件
            if sheet_name:
                df = pd.read_excel(input_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(input_path, sheet_name=0)
            
            # 处理缺失值
            df = df.fillna('')
            
            # 保存为CSV
            df.to_csv(output_path, index=False, encoding=encoding, sep=separator)
            
            return os.path.exists(output_path)
            
        except Exception as e:
            raise ConversionError(f"CSV转换失败: {str(e)}")
    
    def convert_to_pdf(self, input_path: str, output_path: str,
                      page_size: str = 'A4', orientation: str = '纵向',
                      sheet_name: Optional[str] = None) -> bool:
        """
        转换Excel文件为PDF格式
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            page_size: 页面大小 ('A4', 'A3', 'Letter')
            orientation: 页面方向 ('纵向', '横向')
            sheet_name: 工作表名称，None表示第一个工作表
            
        Returns:
            转换是否成功
        """
        try:
            # 读取Excel文件
            if sheet_name:
                df = pd.read_excel(input_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(input_path, sheet_name=0)
            
            # 处理缺失值
            df = df.fillna('')
            
            # 设置页面大小
            if page_size == 'A4':
                pagesize = A4
            elif page_size == 'A3':
                pagesize = A3
            else:  # Letter
                pagesize = letter
            
            # 处理页面方向
            if orientation == '横向':
                pagesize = (pagesize[1], pagesize[0])
            
            # 创建PDF文档
            doc = SimpleDocTemplate(output_path, pagesize=pagesize)
            elements = []
            
            # 获取样式
            styles = getSampleStyleSheet()
            
            # 添加标题
            title = Paragraph(f"数据表格 - {os.path.basename(input_path)}", styles['Title'])
            elements.append(title)
            elements.append(Spacer(1, 12))
            
            # 转换DataFrame为表格数据
            data = []
            
            # 添加表头
            headers = list(df.columns)
            data.append(headers)
            
            # 添加数据行
            for _, row in df.iterrows():
                data.append([str(cell) for cell in row.values])
            
            # 限制显示的行数（避免PDF过大）
            max_rows = 100
            if len(data) > max_rows + 1:  # +1 for header
                data = data[:max_rows + 1]
                # 添加省略提示
                data.append(['...'] * len(headers))
            
            # 创建表格
            table = Table(data)
            
            # 设置表格样式
            table_style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ])
            
            table.setStyle(table_style)
            elements.append(table)
            
            # 生成PDF
            doc.build(elements)
            
            return os.path.exists(output_path)
            
        except Exception as e:
            raise ConversionError(f"PDF转换失败: {str(e)}")
    
    def convert_to_xlsx(self, input_path: str, output_path: str,
                       sheet_name: Optional[str] = None) -> bool:
        """
        转换Excel文件为XLSX格式（主要用于xls到xlsx的转换）
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            sheet_name: 工作表名称，None表示所有工作表
            
        Returns:
            转换是否成功
        """
        try:
            if sheet_name:
                # 转换指定工作表
                df = pd.read_excel(input_path, sheet_name=sheet_name)
                df.to_excel(output_path, index=False, sheet_name=sheet_name)
            else:
                # 转换所有工作表
                excel_file = pd.ExcelFile(input_path)
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    for sheet in excel_file.sheet_names:
                        df = pd.read_excel(input_path, sheet_name=sheet)
                        df.to_excel(writer, sheet_name=sheet, index=False)
            
            return os.path.exists(output_path)
            
        except Exception as e:
            raise ConversionError(f"XLSX转换失败: {str(e)}")
    
    def get_excel_sheets(self, file_path: str) -> List[str]:
        """
        获取Excel文件中的所有工作表名称
        
        Args:
            file_path: Excel文件路径
            
        Returns:
            工作表名称列表
        """
        try:
            excel_file = pd.ExcelFile(file_path)
            return excel_file.sheet_names
        except Exception as e:
            st.error(f"读取工作表失败: {str(e)}")
            return []
    
    def preview_excel_data(self, file_path: str, sheet_name: Optional[str] = None, 
                          max_rows: int = 10) -> Optional[pd.DataFrame]:
        """
        预览Excel文件数据
        
        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称
            max_rows: 最大预览行数
            
        Returns:
            DataFrame或None
        """
        try:
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=max_rows)
            else:
                df = pd.read_excel(file_path, sheet_name=0, nrows=max_rows)
            
            return df
            
        except Exception as e:
            st.error(f"预览数据失败: {str(e)}")
            return None
    
    def validate_input_file(self, file_path: str) -> Dict[str, Any]:
        """
        验证输入文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            验证结果字典
        """
        result = {
            'valid': False,
            'error': '',
            'file_info': {},
            'sheets': []
        }
        
        try:
            if not os.path.exists(file_path):
                result['error'] = '文件不存在'
                return result
            
            # 检查文件扩展名
            file_ext = Path(file_path).suffix.lower().lstrip('.')
            if file_ext not in self.supported_input_formats:
                result['error'] = f'不支持的文件格式: .{file_ext}'
                return result
            
            # 尝试读取文件信息
            file_size = os.path.getsize(file_path)
            result['file_info'] = {
                'name': os.path.basename(file_path),
                'size': file_size,
                'extension': file_ext
            }
            
            # 获取工作表信息
            try:
                sheets = self.get_excel_sheets(file_path)
                result['sheets'] = sheets
            except Exception:
                result['sheets'] = ['Sheet1']  # 默认工作表名
            
            result['valid'] = True
            
        except Exception as e:
            result['error'] = f'文件验证失败: {str(e)}'
        
        return result
    
    def convert_single_file(self, input_path: str, output_path: str, 
                           output_format: str, params: Dict[str, Any],
                           progress_callback=None) -> bool:
        """
        转换单个文件
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            output_format: 输出格式
            params: 转换参数
            progress_callback: 进度回调函数
            
        Returns:
            转换是否成功
        """
        try:
            if progress_callback:
                progress_callback(0.1)  # 开始处理
            
            if output_format == 'csv':
                success = self.convert_to_csv(
                    input_path, output_path,
                    encoding=params.get('encoding', 'utf-8'),
                    separator=params.get('separator', ','),
                    sheet_name=params.get('sheet_name')
                )
            elif output_format == 'pdf':
                success = self.convert_to_pdf(
                    input_path, output_path,
                    page_size=params.get('page_size', 'A4'),
                    orientation=params.get('orientation', '纵向'),
                    sheet_name=params.get('sheet_name')
                )
            elif output_format == 'xlsx':
                success = self.convert_to_xlsx(
                    input_path, output_path,
                    sheet_name=params.get('sheet_name')
                )
            else:
                raise ConversionError(f"不支持的输出格式: {output_format}")
            
            if progress_callback:
                progress_callback(1.0)  # 完成处理
            
            return success
            
        except Exception as e:
            if progress_callback:
                progress_callback(0.0)  # 重置进度
            raise ConversionError(f"文件转换失败: {str(e)}")

    def batch_convert(self, file_configs: List[Dict[str, Any]], 
                     progress_callback=None) -> Dict[str, Any]:
        """
        批量转换文件
        
        Args:
            file_configs: 文件配置列表，每个配置包含输入路径、输出路径、转换参数等
            progress_callback: 进度回调函数
            
        Returns:
            批量转换结果
        """
        results = {
            'total': len(file_configs),
            'success': 0,
            'failed': 0,
            'errors': [],
            'output_files': []
        }
        
        for i, config in enumerate(file_configs):
            try:
                input_path = config['input_path']
                output_path = config['output_path']
                output_format = config['output_format']
                params = config.get('params', {})
                
                # 根据输出格式调用相应的转换方法
                if output_format == 'csv':
                    success = self.convert_to_csv(input_path, output_path, **params)
                elif output_format == 'pdf':
                    success = self.convert_to_pdf(input_path, output_path, **params)
                elif output_format == 'xlsx':
                    success = self.convert_to_xlsx(input_path, output_path, **params)
                else:
                    raise ConversionError(f'不支持的输出格式: {output_format}')
                
                if success:
                    results['success'] += 1
                    results['output_files'].append(output_path)
                else:
                    results['failed'] += 1
                    results['errors'].append(f'{os.path.basename(input_path)}: 转换失败')
                
            except Exception as e:
                results['failed'] += 1
                error_msg = f'{os.path.basename(config["input_path"])}: {str(e)}'
                results['errors'].append(error_msg)
            
            # 更新进度
            if progress_callback:
                progress = (i + 1) / len(file_configs)
                progress_callback(progress)
        
        return results 