# -*- coding: utf-8 -*-
"""
文件下载界面组件
"""

import streamlit as st
import os
import pandas as pd
from typing import List, Dict, Any, Optional
from pathlib import Path
from datetime import datetime

from core.converter import FileConverter, ConversionError
from core.file_handler import FileManager
from core.progress_tracker import ProgressTracker
from config.settings import (
    SUPPORTED_OUTPUT_FORMATS,
    CSV_SEPARATORS,
    CSV_ENCODINGS,
    PDF_ORIENTATIONS,
    PDF_PAGE_SIZES,
    XLSX_OPTIONS,
    XLSX_COMPRESSION_LEVELS
)


class DownloadSection:
    """文件下载界面组件类"""
    
    def __init__(self, file_manager: FileManager, converter: FileConverter):
        """
        初始化下载组件
        
        Args:
            file_manager: 文件管理器实例
            converter: 文件转换器实例
        """
        self.file_manager = file_manager
        self.converter = converter
    
    def render_conversion_config(self, file_info_list: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
        """
        渲染转换配置界面
        
        Args:
            file_info_list: 文件信息列表
            
        Returns:
            转换配置字典或None
        """
        if not file_info_list:
            return None
        
        st.subheader("⚙️ 转换配置")
        
        # 输出格式选择
        output_format = st.selectbox(
            "选择输出格式:",
            SUPPORTED_OUTPUT_FORMATS,
            format_func=lambda x: {
                'csv': 'CSV (逗号分隔值)',
                'pdf': 'PDF (便携式文档)',
                'xlsx': 'XLSX (Excel工作簿)'
            }[x]
        )
        
        # 根据输出格式显示相应配置
        format_config = {}
        
        if output_format == 'csv':
            format_config = self._render_csv_config()
        elif output_format == 'pdf':
            format_config = self._render_pdf_config()
        elif output_format == 'xlsx':
            format_config = self._render_xlsx_config()
        
        # 工作表处理选项
        sheet_config = self._render_sheet_config(file_info_list)
        
        return {
            'output_format': output_format,
            'format_config': format_config,
            'sheet_config': sheet_config
        }
    
    def _render_csv_config(self) -> Dict[str, Any]:
        """渲染CSV格式配置"""
        st.write("**CSV配置选项:**")
        
        col1, col2 = st.columns(2)
        
        with col1:
            encoding = st.selectbox(
                "文件编码:",
                CSV_ENCODINGS,
                index=0  # 默认UTF-8
            )
        
        with col2:
            separator_name = st.selectbox(
                "分隔符:",
                list(CSV_SEPARATORS.keys()),
                index=0  # 默认逗号
            )
            separator = CSV_SEPARATORS[separator_name]
        
        return {
            'encoding': encoding,
            'separator': separator
        }
    
    def _render_pdf_config(self) -> Dict[str, Any]:
        """渲染PDF格式配置"""
        st.write("**PDF配置选项:**")
        
        col1, col2 = st.columns(2)
        
        with col1:
            page_size = st.selectbox(
                "页面大小:",
                list(PDF_PAGE_SIZES.keys()),
                index=0  # 默认A4
            )
        
        with col2:
            orientation = st.selectbox(
                "页面方向:",
                PDF_ORIENTATIONS,
                index=0  # 默认纵向
            )
        
        return {
            'page_size': page_size,
            'orientation': orientation
        }
    
    def _render_xlsx_config(self) -> Dict[str, Any]:
        """渲染XLSX格式配置"""
        st.write("**XLSX配置选项:**")
        
        # 基础设置
        col1, col2 = st.columns(2)
        
        with col1:
            preserve_format = st.checkbox(
                "保留原始格式",
                value=XLSX_OPTIONS['保留原始格式'],
                help="保持原始Excel文件的格式和样式"
            )
            
            preserve_formulas = st.checkbox(
                "保留公式",
                value=XLSX_OPTIONS['保留公式'],
                help="保留Excel中的公式，否则转换为计算值"
            )
        
        with col2:
            protect_sheets = st.checkbox(
                "工作表保护",
                value=XLSX_OPTIONS['工作表保护'],
                help="对转换后的工作表进行保护"
            )
            
            memory_optimize = st.checkbox(
                "内存优化模式",
                value=XLSX_OPTIONS['内存优化模式'],
                help="适用于大文件处理，可能影响处理速度"
            )
        
        # 高级设置
        with st.expander("高级设置", expanded=False):
            compression_name = st.selectbox(
                "压缩级别:",
                list(XLSX_COMPRESSION_LEVELS.keys()),
                index=2,  # 默认标准压缩
                help="更高压缩级别会减小文件大小但增加处理时间"
            )
            compression_level = XLSX_COMPRESSION_LEVELS[compression_name]
            
            max_rows = st.number_input(
                "最大行数限制:",
                min_value=1000,
                max_value=1000000,
                value=100000,
                step=1000,
                help="限制处理的最大行数，避免内存溢出"
            )
        
        return {
            'preserve_format': preserve_format,
            'preserve_formulas': preserve_formulas,
            'protect_sheets': protect_sheets,
            'memory_optimize': memory_optimize,
            'compression_level': compression_level,
            'max_rows': max_rows
        }
    
    def _render_sheet_config(self, file_info_list: List[Dict[str, Any]]) -> Dict[str, Any]:
        """渲染工作表配置"""
        st.write("**工作表处理:**")
        
        # 检查是否有多工作表文件
        has_multiple_sheets = any(len(info['sheets']) > 1 for info in file_info_list)
        
        if has_multiple_sheets:
            sheet_mode = st.radio(
                "工作表处理方式:",
                ['所有工作表', '仅第一个工作表', '选择特定工作表'],
                index=0
            )
            
            specific_sheets = {}
            if sheet_mode == '选择特定工作表':
                st.write("为每个文件选择要转换的工作表:")
                for info in file_info_list:
                    if len(info['sheets']) > 1:
                        selected_sheet = st.selectbox(
                            f"{info['name']}:",
                            info['sheets'],
                            key=f"sheet_select_{info['name']}"
                        )
                        specific_sheets[info['name']] = selected_sheet
                    else:
                        specific_sheets[info['name']] = info['sheets'][0]
        else:
            sheet_mode = '所有工作表'
            specific_sheets = {}
        
        return {
            'mode': sheet_mode,
            'specific_sheets': specific_sheets
        }
    
    def render_conversion_action(self, file_info_list: List[Dict[str, Any]], 
                               conversion_config: Dict[str, Any]) -> Optional[tuple]:
        """
        渲染转换执行界面
        
        Args:
            file_info_list: 文件信息列表
            conversion_config: 转换配置
            
        Returns:
            (ZIP文件路径, 输出目录)元组或None
        """
        if not file_info_list or not conversion_config:
            return None
        
        st.subheader("🚀 开始转换")
        
        # 转换摘要
        col1, col2 = st.columns([2, 1])
        
        with col1:
            st.write(f"**待转换文件:** {len(file_info_list)} 个")
            st.write(f"**输出格式:** {conversion_config['output_format'].upper()}")
            st.write(f"**工作表模式:** {conversion_config['sheet_config']['mode']}")
        
        with col2:
            if st.button("🔄 开始批量转换", type="primary", use_container_width=True):
                st.info("开始执行转换...")
                result = self._execute_batch_conversion(file_info_list, conversion_config)
                if result:
                    st.success("转换完成！")
                else:
                    st.error("转换失败！")
                return result
        
        return None
    
    def _execute_batch_conversion(self, file_info_list: List[Dict[str, Any]], 
                                 conversion_config: Dict[str, Any]) -> Optional[tuple]:
        """
        执行批量转换
        
        Args:
            file_info_list: 文件信息列表
            conversion_config: 转换配置
            
        Returns:
            (ZIP文件路径, 输出目录)元组或None
        """
        try:
            # 创建输出目录
            output_dir = self.file_manager.create_temp_directory("converted")
            
            # 准备转换配置
            file_configs = self._prepare_conversion_configs(
                file_info_list, 
                conversion_config, 
                output_dir
            )
            
            if not file_configs:
                st.error("没有可转换的文件配置")
                return None
            
            # 创建进度跟踪器
            progress_tracker = ProgressTracker(
                total_files=len(file_configs),
                enable_time_estimation=True
            )
            
            # 初始化文件进度
            progress_tracker.initialize_files(file_info_list)
            progress_tracker.setup_ui()
            progress_tracker.start_processing()
            
            # 执行增强的批量转换
            results = self._execute_enhanced_batch_conversion(
                file_configs, progress_tracker
            )
            
            # 显示转换结果
            self._display_conversion_results(results)
            
            if results['success'] > 0:
                # 创建ZIP文件
                progress_tracker.update_status("打包转换结果...")
                zip_path = self.file_manager.zip_converted_files(
                    results['output_files'],
                    f"converted_{conversion_config['output_format']}_files.zip"
                )
                
                # 完成进度跟踪
                stats = progress_tracker.finish(success=True)
                
                # 保存输出目录到结果中
                results['output_dir'] = output_dir
                results['stats'] = stats
                
                return zip_path, output_dir
            
            return None
            
        except Exception as e:
            st.error(f"批量转换失败: {str(e)}")
            return None
    
    def _execute_enhanced_batch_conversion(self, file_configs: List[Dict[str, Any]], 
                                         progress_tracker: ProgressTracker) -> Dict[str, Any]:
        """
        执行增强的批量转换（带详细进度跟踪）
        
        Args:
            file_configs: 文件配置列表
            progress_tracker: 进度跟踪器
            
        Returns:
            转换结果字典
        """
        results = {
            'total': len(file_configs),
            'success': 0,
            'failed': 0,
            'output_files': [],
            'errors': []
        }
        
        try:
            for i, config in enumerate(file_configs):
                # 开始处理文件
                progress_tracker.start_file(i)
                
                try:
                    # 创建文件级别的进度回调
                    file_callback = progress_tracker.create_file_callback(i)
                    
                    # 执行单个文件转换
                    success = self.converter.convert_single_file(
                        config['input_path'],
                        config['output_path'],
                        config['output_format'],
                        config['params'],
                        progress_callback=file_callback
                    )
                    
                    if success:
                        results['success'] += 1
                        results['output_files'].append(config['output_path'])
                        progress_tracker.complete_file(i, success=True)
                    else:
                        results['failed'] += 1
                        error_msg = f"转换失败: {os.path.basename(config['input_path'])}"
                        results['errors'].append(error_msg)
                        progress_tracker.complete_file(i, success=False, error_message=error_msg)
                        
                except Exception as e:
                    results['failed'] += 1
                    error_msg = f"处理文件时发生错误: {os.path.basename(config['input_path'])}: {str(e)}"
                    results['errors'].append(error_msg)
                    progress_tracker.complete_file(i, success=False, error_message=str(e))
                    
        except Exception as e:
            progress_tracker.finish(success=False)
            raise e
        
        return results
    
    def _prepare_conversion_configs(self, file_info_list: List[Dict[str, Any]], 
                                  conversion_config: Dict[str, Any], 
                                  output_dir: str) -> List[Dict[str, Any]]:
        """准备转换配置列表"""
        configs = []
        
        output_format = conversion_config['output_format']
        format_config = conversion_config['format_config']
        sheet_config = conversion_config['sheet_config']
        
        for file_info in file_info_list:
            input_path = file_info['path']
            file_name = Path(file_info['name']).stem
            
            # 根据工作表处理模式准备配置
            if sheet_config['mode'] == '所有工作表':
                if len(file_info['sheets']) > 1 and output_format in ['csv', 'pdf']:
                    # 多工作表文件，每个工作表单独转换
                    for sheet_name in file_info['sheets']:
                        output_name = f"{file_name}_{sheet_name}.{output_format}"
                        output_path = os.path.join(output_dir, output_name)
                        
                        config = {
                            'input_path': input_path,
                            'output_path': output_path,
                            'output_format': output_format,
                            'params': {**format_config, 'sheet_name': sheet_name}
                        }
                        configs.append(config)
                else:
                    # 单工作表或XLSX格式
                    output_name = f"{file_name}.{output_format}"
                    output_path = os.path.join(output_dir, output_name)
                    
                    config = {
                        'input_path': input_path,
                        'output_path': output_path,
                        'output_format': output_format,
                        'params': format_config
                    }
                    configs.append(config)
                    
            elif sheet_config['mode'] == '仅第一个工作表':
                output_name = f"{file_name}.{output_format}"
                output_path = os.path.join(output_dir, output_name)
                
                config = {
                    'input_path': input_path,
                    'output_path': output_path,
                    'output_format': output_format,
                    'params': {**format_config, 'sheet_name': file_info['sheets'][0]}
                }
                configs.append(config)
                
            elif sheet_config['mode'] == '选择特定工作表':
                selected_sheet = sheet_config['specific_sheets'].get(file_info['name'])
                if selected_sheet:
                    output_name = f"{file_name}_{selected_sheet}.{output_format}"
                    output_path = os.path.join(output_dir, output_name)
                    
                    config = {
                        'input_path': input_path,
                        'output_path': output_path,
                        'output_format': output_format,
                        'params': {**format_config, 'sheet_name': selected_sheet}
                    }
                    configs.append(config)
        
        return configs
    
    def _display_conversion_results(self, results: Dict[str, Any]):
        """显示转换结果"""
        # 使用简单的文本显示而不是列布局
        st.write("**转换结果:**")
        st.write(f"总文件数: {results['total']} | 成功转换: {results['success']} | 转换失败: {results['failed']}")
        
        # 显示错误信息
        if results['errors']:
            st.error("转换过程中遇到以下错误:")
            for error in results['errors']:
                st.write(f"• {error}")
    
    def render_download_area(self, zip_path: str, output_dir: Optional[str] = None):
        """
        渲染下载区域
        
        Args:
            zip_path: ZIP文件路径
            output_dir: 输出目录路径
        """
        if not zip_path or not os.path.exists(zip_path):
            return
        
        st.subheader("📥 下载转换结果")
        
        # 下载选项标签页
        tab1, tab2 = st.tabs(["📦 批量下载", "📄 单文件下载"])
        
        with tab1:
            # 批量下载ZIP文件
            file_info = self.file_manager.get_file_info(zip_path)
            
            st.write(f"**文件名:** {file_info['name']}")
            st.write(f"**文件大小:** {file_info['size_formatted']}")
            st.write("**包含:** 所有转换后的文件")
            
            # 读取文件内容用于下载
            with open(zip_path, 'rb') as f:
                file_data = f.read()
            
            st.download_button(
                label="📁 下载ZIP文件",
                data=file_data,
                file_name=file_info['name'],
                mime="application/zip",
                type="primary",
                use_container_width=True
            )
        
        with tab2:
            # 单个文件下载列表
            if output_dir:
                self._render_individual_downloads(output_dir)
            else:
                st.info("需要输出目录信息才能显示单文件下载")
        
        st.success("转换完成！选择下载方式获取文件。")
    
    def _render_individual_downloads(self, output_dir: str) -> None:
        """
        渲染单个文件下载区域
        
        Args:
            output_dir: 输出目录路径
        """
        converted_files = self.file_manager.list_converted_files(output_dir)
        
        if not converted_files:
            st.info("没有找到转换完成的文件")
            return
        
        st.write(f"共 {len(converted_files)} 个文件可供下载:")
        
        # 创建文件列表 - 简化显示避免嵌套列
        for i, file_info in enumerate(converted_files):
            with st.container():
                # 文件名和图标
                icon = self._get_file_icon(file_info['extension'])
                format_type = file_info['extension'].upper().replace('.', '')
                
                st.write(f"{icon} **{file_info['download_name']}** `{format_type}` ({file_info['size_formatted']})")
                
                # 操作按钮区域
                button_col1, button_col2, button_col3 = st.columns([1, 1, 4])
                
                with button_col1:
                    # 预览按钮（仅CSV和XLSX）
                    if file_info['extension'] in ['.csv', '.xlsx']:
                        if st.button("👀 预览", key=f"download_preview_{i}"):
                            self._show_file_preview(file_info['full_path'])
                
                with button_col2:
                    # 下载按钮
                    file_data = self.file_manager.generate_individual_download_data(file_info['full_path'])
                    if file_data:
                        st.download_button(
                            label="⬇️ 下载",
                            data=file_data,
                            file_name=file_info['download_name'],
                            mime=file_info['mime_type'],
                            key=f"download_{i}"
                        )
                
                # 分隔线
                if i < len(converted_files) - 1:
                    st.divider()
    
    def _get_file_icon(self, extension: str) -> str:
        """获取文件图标"""
        icons = {
            '.xlsx': '📊',
            '.csv': '📋',
            '.pdf': '📄',
            '.xls': '📈'
        }
        return icons.get(extension, '📁')
    
    def _show_file_preview(self, file_path: str) -> None:
        """显示文件预览"""
        try:
            if file_path.endswith('.csv'):
                # 尝试多种编码读取CSV文件
                encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb2312']
                df = None
                
                for encoding in encodings:
                    try:
                        df = pd.read_csv(file_path, nrows=5, encoding=encoding)
                        break  # 成功读取就退出循环
                    except UnicodeDecodeError:
                        continue
                    except Exception:
                        continue
                
                if df is None:
                    st.error("无法以正确编码读取CSV文件")
                    return
                    
            elif file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path, nrows=5, engine='openpyxl')
            else:
                st.error("不支持预览此文件类型")
                return
            
            with st.expander(f"预览: {os.path.basename(file_path)}", expanded=True):
                st.dataframe(df, use_container_width=True)
                if len(df) >= 5:
                    st.caption("显示前5行，完整内容请下载文件查看")
                    
        except Exception as e:
            st.error(f"预览失败: {str(e)}") 