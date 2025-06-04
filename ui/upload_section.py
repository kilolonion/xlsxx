# -*- coding: utf-8 -*-
"""
文件上传界面组件
"""

import streamlit as st
from typing import List, Optional, Dict, Any
import pandas as pd

from core.converter import FileConverter
from core.file_handler import FileManager
from config.settings import (
    MAX_FILE_SIZE, 
    SUPPORTED_INPUT_FORMATS, 
    UPLOAD_HELP_TEXT,
    MAX_PREVIEW_ROWS
)


class UploadSection:
    """文件上传界面组件类"""
    
    def __init__(self, file_manager: FileManager, converter: FileConverter):
        """
        初始化上传组件
        
        Args:
            file_manager: 文件管理器实例
            converter: 文件转换器实例
        """
        self.file_manager = file_manager
        self.converter = converter
    
    def render_upload_area(self) -> List[Dict[str, Any]]:
        """
        渲染文件上传区域
        
        Returns:
            上传文件信息列表
        """
        st.subheader("📁 文件上传")
        
        # 创建文件上传组件
        uploaded_files = st.file_uploader(
            label="选择Excel文件进行转换",
            type=SUPPORTED_INPUT_FORMATS,
            accept_multiple_files=True,
            help=UPLOAD_HELP_TEXT
        )
        
        uploaded_file_info = []
        
        if uploaded_files:
            st.success(f"已选择 {len(uploaded_files)} 个文件")
            
            # 处理每个上传的文件
            for i, uploaded_file in enumerate(uploaded_files):
                file_info = self._process_uploaded_file(uploaded_file, i)
                if file_info:
                    uploaded_file_info.append(file_info)
        
        return uploaded_file_info
    
    def _process_uploaded_file(self, uploaded_file, index: int) -> Optional[Dict[str, Any]]:
        """
        处理单个上传文件
        
        Args:
            uploaded_file: Streamlit上传文件对象
            index: 文件索引
            
        Returns:
            文件信息字典或None
        """
        try:
            # 显示文件基本信息
            with st.expander(f"📄 {uploaded_file.name}", expanded=True):
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    st.write(f"**文件名:** {uploaded_file.name}")
                    st.write(f"**文件大小:** {self._format_file_size(uploaded_file.size)}")
                    st.write(f"**文件类型:** .{uploaded_file.name.split('.')[-1].upper()}")
                
                # 保存文件到临时目录
                saved_path = self.file_manager.save_uploaded_file(uploaded_file)
                
                if not saved_path:
                    st.error("文件保存失败")
                    return None
                
                # 验证文件
                validation_result = self.converter.validate_input_file(saved_path)
                
                if not validation_result['valid']:
                    st.error(f"文件验证失败: {validation_result['error']}")
                    return None
                
                # 获取工作表信息
                sheets = validation_result['sheets']
                
                with col2:
                    st.write(f"**工作表数量:** {len(sheets)}")
                    if len(sheets) > 1:
                        st.write("**工作表列表:**")
                        for sheet in sheets:
                            st.write(f"• {sheet}")
                
                # 数据预览选项
                if st.checkbox(f"预览数据", key=f"upload_preview_{index}"):
                    self._render_data_preview(saved_path, sheets, index)
                
                # 返回文件信息
                return {
                    'name': uploaded_file.name,
                    'path': saved_path,
                    'size': uploaded_file.size,
                    'sheets': sheets,
                    'validation': validation_result
                }
                
        except Exception as e:
            st.error(f"处理文件 {uploaded_file.name} 时出错: {str(e)}")
            return None
    
    def _render_data_preview(self, file_path: str, sheets: List[str], file_index: int):
        """
        渲染数据预览
        
        Args:
            file_path: 文件路径
            sheets: 工作表列表
            file_index: 文件索引
        """
        try:
            # 选择工作表
            if len(sheets) > 1:
                selected_sheet = st.selectbox(
                    "选择要预览的工作表:",
                    sheets,
                    key=f"sheet_select_{file_index}"
                )
            else:
                selected_sheet = sheets[0] if sheets else None
            
            if selected_sheet:
                # 预览数据
                df = self.converter.preview_excel_data(
                    file_path, 
                    selected_sheet, 
                    MAX_PREVIEW_ROWS
                )
                
                if df is not None:
                    st.write(f"**预览 {selected_sheet} (前{len(df)}行):**")
                    st.dataframe(df, use_container_width=True)
                    
                    # 显示基本统计信息
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("行数", len(df))
                    with col2:
                        st.metric("列数", len(df.columns))
                    with col3:
                        st.metric("空值数", df.isnull().sum().sum())
                else:
                    st.warning("无法预览数据")
                    
        except Exception as e:
            st.error(f"预览数据时出错: {str(e)}")
    
    def render_batch_info(self, file_info_list: List[Dict[str, Any]]):
        """
        渲染批量处理信息摘要
        
        Args:
            file_info_list: 文件信息列表
        """
        if not file_info_list:
            return
        
        st.subheader("📊 批量处理摘要")
        
        # 统计信息
        total_files = len(file_info_list)
        total_size = sum(info['size'] for info in file_info_list)
        total_sheets = sum(len(info['sheets']) for info in file_info_list)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric(
                label="文件总数",
                value=total_files,
                help="待转换的文件数量"
            )
        
        with col2:
            st.metric(
                label="总大小",
                value=self._format_file_size(total_size),
                help="所有文件的总大小"
            )
        
        with col3:
            st.metric(
                label="工作表总数",
                value=total_sheets,
                help="所有文件包含的工作表总数"
            )
        
        # 文件列表
        with st.expander("查看文件详情", expanded=False):
            for i, info in enumerate(file_info_list, 1):
                st.write(f"**{i}. {info['name']}**")
                st.write(f"   - 大小: {self._format_file_size(info['size'])}")
                st.write(f"   - 工作表: {', '.join(info['sheets'])}")
    
    def _format_file_size(self, size_bytes: int) -> str:
        """格式化文件大小显示"""
        if size_bytes == 0:
            return "0B"
        
        size_names = ["B", "KB", "MB", "GB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
        
        return f"{size_bytes:.1f}{size_names[i]}" 