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
    
    def _read_excel_with_fallback(self, input_path: str, sheet_name: Optional[str] = None, 
                                 nrows: Optional[int] = None) -> pd.DataFrame:
        """
        使用备用引擎读取Excel文件
        
        Args:
            input_path: 输入文件路径
            sheet_name: 工作表名称，None表示第一个工作表或使用0
            nrows: 最大读取行数
            
        Returns:
            DataFrame
            
        Raises:
            ConversionError: 当所有引擎都失败时
        """
        file_ext = Path(input_path).suffix.lower()
        
        # 准备读取参数
        read_params = {}
        if sheet_name is not None:
            read_params['sheet_name'] = sheet_name
        else:
            read_params['sheet_name'] = 0  # 默认第一个工作表
            
        if nrows is not None:
            read_params['nrows'] = nrows
        
        # 根据文件扩展名确定引擎优先级
        if file_ext == '.xlsx':
            engines = ['openpyxl']
        else:  # .xls
            engines = ['xlrd', 'openpyxl']
        
        last_error = None
        
        for engine in engines:
            try:
                if engine == 'xlrd':
                    # xlrd引擎对老版本XLS支持更好，但需要特殊编码处理
                    import xlrd
                    
                    # 打开工作簿，启用格式化信息以获得更好的编码支持
                    workbook = xlrd.open_workbook(input_path, encoding_override='utf-8')
                    
                    # 确定要读取的工作表
                    if sheet_name is not None and isinstance(sheet_name, str):
                        if sheet_name in workbook.sheet_names():
                            worksheet = workbook.sheet_by_name(sheet_name)
                        else:
                            # 如果指定的工作表不存在，使用第一个
                            worksheet = workbook.sheet_by_index(0)
                    else:
                        # 使用索引或第一个工作表
                        sheet_index = sheet_name if isinstance(sheet_name, int) else 0
                        worksheet = workbook.sheet_by_index(sheet_index)
                    
                    # 读取数据并确保正确编码
                    data = []
                    max_rows = nrows if nrows else worksheet.nrows
                    actual_rows = min(max_rows, worksheet.nrows)
                    
                    for row_idx in range(actual_rows):
                        row_data = []
                        for col_idx in range(worksheet.ncols):
                            cell = worksheet.cell(row_idx, col_idx)
                            cell_value = cell.value
                            
                            # 处理不同类型的单元格值
                            if cell.ctype == xlrd.XL_CELL_TEXT:
                                # 文本类型，确保UTF-8编码
                                if isinstance(cell_value, bytes):
                                    try:
                                        cell_value = cell_value.decode('utf-8')
                                    except UnicodeDecodeError:
                                        try:
                                            cell_value = cell_value.decode('gbk')
                                        except UnicodeDecodeError:
                                            cell_value = cell_value.decode('latin-1')
                                elif isinstance(cell_value, str):
                                    # 已经是字符串，保持原样
                                    pass
                            elif cell.ctype == xlrd.XL_CELL_NUMBER:
                                # 数字类型
                                if cell_value == int(cell_value):
                                    cell_value = int(cell_value)
                            elif cell.ctype == xlrd.XL_CELL_DATE:
                                # 日期类型
                                if xlrd.xldate.xldate_is_date(cell_value, workbook.datemode):
                                    cell_value = xlrd.xldate.xldate_as_datetime(cell_value, workbook.datemode)
                            elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                                # 布尔类型
                                cell_value = bool(cell_value)
                            elif cell.ctype == xlrd.XL_CELL_EMPTY:
                                # 空单元格
                                cell_value = ''
                            
                            row_data.append(cell_value)
                        
                        data.append(row_data)
                    
                    # 创建DataFrame
                    if data:
                        # 第一行通常是表头
                        if len(data) > 1:
                            df = pd.DataFrame(data[1:], columns=data[0])
                        else:
                            df = pd.DataFrame(data)
                    else:
                        df = pd.DataFrame()
                    
                    # 成功读取，返回结果
                    return df
                    
                else:
                    # openpyxl引擎
                    df = pd.read_excel(input_path, engine='openpyxl', **read_params)
                
                # 成功读取，返回结果
                return df
                
            except Exception as e:
                last_error = e
                error_msg = str(e).lower()
                
                # 记录具体的错误信息
                if 'no valid workbook part' in error_msg:
                    st.warning(f"引擎 {engine} 无法读取文件（格式不兼容），尝试备用引擎...")
                elif 'xlrd' in error_msg and 'xlsx' in error_msg:
                    st.warning(f"XLS引擎不支持XLSX格式，尝试XLSX引擎...")
                elif 'file is not a zip file' in error_msg:
                    st.warning(f"文件可能是老版本XLS格式，尝试专用引擎...")
                else:
                    st.warning(f"引擎 {engine} 读取失败: {str(e)[:100]}")
                
                continue
        
        # 所有引擎都失败了
        if last_error:
            raise ConversionError(f"无法读取Excel文件（已尝试所有可用引擎）: {str(last_error)}")
        else:
            raise ConversionError("无法读取Excel文件：没有可用的读取引擎")
    
    def convert_to_csv(self, input_path: str, output_path: str, 
                      encoding: str = 'utf-8-sig', separator: str = ',',
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
            # 使用增强的读取方法
            df = self._read_excel_with_fallback(input_path, sheet_name)
            
            # 处理缺失值
            df = df.fillna('')
            
            # 确保DataFrame中的所有文本数据都是正确编码的字符串
            for col in df.columns:
                if df[col].dtype == 'object':  # 通常是字符串类型
                    df[col] = df[col].astype(str)
            
            # 确保使用UTF-8 BOM编码保存，这样Excel可以正确识别中文
            if encoding == 'utf-8':
                encoding = 'utf-8-sig'
            
            # 保存为CSV，显式指定错误处理方式
            df.to_csv(output_path, index=False, encoding=encoding, sep=separator, 
                     errors='replace')  # 使用replace处理编码错误
            
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
            # 使用增强的读取方法
            df = self._read_excel_with_fallback(input_path, sheet_name)
            
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
                df = self._read_excel_with_fallback(input_path, sheet_name)
                df.to_excel(output_path, index=False, sheet_name=sheet_name, engine='openpyxl')
            else:
                # 转换所有工作表
                try:
                    # 尝试获取所有工作表
                    all_sheets = self.get_excel_sheets(input_path)
                    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                        for sheet in all_sheets:
                            df = self._read_excel_with_fallback(input_path, sheet)
                            df.to_excel(writer, sheet_name=sheet, index=False)
                except Exception:
                    # 如果失败，只转换第一个工作表
                    df = self._read_excel_with_fallback(input_path, None)
                    df.to_excel(output_path, index=False, engine='openpyxl')
            
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
        file_ext = Path(file_path).suffix.lower()
        
        # 根据文件类型选择引擎
        if file_ext == '.xlsx':
            engines = ['openpyxl']
        else:  # .xls
            engines = ['xlrd', 'openpyxl']
        
        for engine in engines:
            try:
                if engine == 'xlrd':
                    import xlrd
                    
                    # 打开工作簿，处理编码
                    workbook = xlrd.open_workbook(file_path, encoding_override='utf-8')
                    
                    # 获取工作表名称并确保正确编码
                    sheet_names = []
                    for name in workbook.sheet_names():
                        if isinstance(name, bytes):
                            # 如果是字节类型，尝试解码
                            try:
                                name = name.decode('utf-8')
                            except UnicodeDecodeError:
                                try:
                                    name = name.decode('gbk')
                                except UnicodeDecodeError:
                                    name = name.decode('latin-1')
                        sheet_names.append(name)
                    
                    return sheet_names
                    
                else:
                    excel_file = pd.ExcelFile(file_path, engine='openpyxl')
                    return excel_file.sheet_names
                
            except Exception as e:
                continue
        
        # 如果所有引擎都失败，返回默认工作表名
        st.warning(f"无法读取工作表信息，使用默认设置")
        return ['Sheet1']
    
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
            df = self._read_excel_with_fallback(file_path, sheet_name, nrows=max_rows)
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
            'sheets': [],
            'warnings': []
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
            
            # 尝试读取文件内容进行验证
            try:
                # 使用备用引擎验证文件可读性
                test_df = self._read_excel_with_fallback(file_path, None, nrows=1)
                if test_df is not None and len(test_df.columns) > 0:
                    result['valid'] = True
                    
                    # 获取工作表信息
                    try:
                        sheets = self.get_excel_sheets(file_path)
                        result['sheets'] = sheets
                        
                        # 检查是否有空工作表
                        if not sheets:
                            result['warnings'].append('未找到有效工作表')
                            
                    except Exception as e:
                        result['warnings'].append(f'读取工作表信息时出现问题: {str(e)[:50]}')
                        result['sheets'] = ['Sheet1']  # 使用默认工作表名
                else:
                    result['error'] = '文件内容为空或格式无效'
                    
            except Exception as e:
                error_msg = str(e).lower()
                
                if 'no valid workbook part' in error_msg:
                    result['error'] = '文件格式损坏或不是有效的Excel文件。请检查文件是否完整，或尝试在Excel中重新保存。'
                elif 'file is not a zip file' in error_msg:
                    result['error'] = '文件可能是损坏的Excel文件或不是标准格式。建议用Excel打开并重新保存。'
                elif 'permission' in error_msg:
                    result['error'] = '文件被其他程序占用或没有读取权限，请关闭文件后重试。'
                elif 'corrupt' in error_msg or 'damaged' in error_msg:
                    result['error'] = '文件已损坏，无法读取。请检查文件完整性。'
                else:
                    result['error'] = f'文件验证失败: {str(e)}'
                    
                # 为特定错误提供建议
                if result['error']:
                    if 'excel' in result['error'].lower():
                        result['warnings'].append('建议: 用Microsoft Excel打开文件，然后另存为新的Excel文件')
                    
        except Exception as e:
            result['error'] = f'文件验证过程中发生错误: {str(e)}'
        
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