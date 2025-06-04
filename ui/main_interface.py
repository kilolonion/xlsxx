# -*- coding: utf-8 -*-
"""
主界面组件
整合所有UI组件并管理应用状态
"""

import streamlit as st
from typing import Dict, Any, Optional

from core.converter import FileConverter
from core.file_handler import FileManager
from .upload_section import UploadSection
from .download_section import DownloadSection
from config.settings import (
    APP_TITLE, 
    APP_ICON, 
    APP_DESCRIPTION,
    MAX_FILES_PER_BATCH
)


class MainInterface:
    """主界面组件类"""
    
    def __init__(self):
        """初始化主界面"""
        self._init_session_state()
        self._init_components()
    
    def _init_session_state(self):
        """初始化会话状态"""
        if 'file_manager' not in st.session_state:
            st.session_state.file_manager = FileManager()
        
        if 'converter' not in st.session_state:
            st.session_state.converter = FileConverter()
        
        if 'uploaded_files' not in st.session_state:
            st.session_state.uploaded_files = []
        
        if 'conversion_config' not in st.session_state:
            st.session_state.conversion_config = None
        
        if 'zip_path' not in st.session_state:
            st.session_state.zip_path = None
            
        if 'output_dir' not in st.session_state:
            st.session_state.output_dir = None
    
    def _init_components(self):
        """初始化UI组件"""
        self.upload_section = UploadSection(
            st.session_state.file_manager,
            st.session_state.converter
        )
        
        self.download_section = DownloadSection(
            st.session_state.file_manager,
            st.session_state.converter
        )
    
    def render(self):
        """渲染主界面"""
        self._render_header()
        self._render_sidebar()
        self._render_main_content()
        self._render_footer()
    
    def _render_header(self):
        """渲染页面头部"""
        st.set_page_config(
            page_title=APP_TITLE,
            page_icon=APP_ICON,
            layout="wide",
            initial_sidebar_state="expanded"
        )

        # 响应式布局，移动端自动纵向排列
        st.markdown(
            """
            <style>
                @media screen and (max-width: 768px) {
                    .stApp [data-testid="stHorizontalBlock"] {
                        flex-direction: column;
                    }
                    .stApp .block-container {
                        padding: 1rem;
                    }
                }
            </style>
            """,
            unsafe_allow_html=True,
        )
        
        # 主标题
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.title(f"{APP_ICON} {APP_TITLE}")
            st.markdown(f"<p style='text-align: center; color: #666;'>{APP_DESCRIPTION}</p>", 
                       unsafe_allow_html=True)
        
        st.divider()
    
    def _render_sidebar(self):
        """渲染侧边栏"""
        with st.sidebar:
            st.header("🛠️ 工具信息")
            
            # 功能特性
            st.subheader("✨ 主要功能")
            st.markdown("""
            - 📊 支持XLS、XLSX文件格式
            - 🔄 批量转换多个文件
            - 📝 输出CSV、PDF、XLSX格式
            - 🎯 灵活的工作表处理选项
            - 📁 一键下载ZIP打包结果
            """)
            
            # 使用说明
            st.subheader("📖 使用说明")
            st.markdown("""
            1. **上传文件**: 选择或拖拽Excel文件
            2. **预览数据**: 可选择预览文件内容
            3. **配置转换**: 选择输出格式和参数
            4. **开始转换**: 执行批量转换操作
            5. **下载结果**: 下载ZIP打包文件
            """)
            
            # 限制说明
            st.subheader("⚠️ 使用限制")
            st.markdown(f"""
            - 单个文件最大: **50MB**
            - 批量处理最多: **{MAX_FILES_PER_BATCH}个文件**
            - 临时文件自动清理: **24小时**
            """)
            
            # 清理操作
            st.subheader("🧹 清理操作")
            if st.button("清理临时文件", type="secondary"):
                self._cleanup_temp_files()
            
            # 重置应用
            if st.button("重置应用", type="secondary"):
                self._reset_application()
    
    def _render_main_content(self):
        """渲染主要内容区域"""
        # 创建两列布局
        col1, col2 = st.columns([1, 1])
        
        with col1:
            # 文件上传区域
            uploaded_files = self.upload_section.render_upload_area()
            
            # 更新会话状态（仅当文件实际变动时）
            def _files_changed(new_files, old_files):
                if len(new_files) != len(old_files):
                    return True
                # 比较文件名列表
                new_names = [f.get('name') for f in new_files]
                old_names = [f.get('name') for f in old_files]
                return new_names != old_names
            
            if _files_changed(uploaded_files, st.session_state.uploaded_files):
                st.session_state.uploaded_files = uploaded_files
                st.session_state.conversion_config = None
                st.session_state.zip_path = None
                st.session_state.output_dir = None
                st.rerun()
        
        with col2:
            # 转换配置和执行区域
            if st.session_state.uploaded_files:
                conversion_config = self.download_section.render_conversion_config(
                    st.session_state.uploaded_files
                )
                
                if conversion_config:
                    st.session_state.conversion_config = conversion_config
                    
                    # 转换执行
                    result = self.download_section.render_conversion_action(
                        st.session_state.uploaded_files,
                        conversion_config
                    )
                    
                    if result:
                        if isinstance(result, tuple) and len(result) == 2:
                            # 新格式：(zip_path, output_dir)
                            zip_path, output_dir = result
                            st.session_state.zip_path = zip_path
                            st.session_state.output_dir = output_dir
                        else:
                            # 旧格式：只有zip_path
                            st.session_state.zip_path = result
                            st.session_state.output_dir = None
                        # 直接渲染下载区域，避免闪烁
                        st.rerun()
            else:
                # 显示欢迎信息
                self._render_welcome_info()
        
        # 批量处理摘要（跨列显示）
        if st.session_state.uploaded_files:
            st.divider()
            self.upload_section.render_batch_info(st.session_state.uploaded_files)

        # 下载区域（跨列显示）
        if st.session_state.zip_path:
            st.divider()
            self.download_section.render_download_area(
                st.session_state.zip_path,
                st.session_state.output_dir
            )
    
    def _render_welcome_info(self):
        """渲染欢迎信息"""
        st.info("👈 请先在左侧上传Excel文件以开始转换")
        
        st.subheader("📋 支持的转换类型")
        
        # 创建转换矩阵
        conversion_matrix = {
            "XLS → CSV": "✅ 支持所有工作表分别转换",
            "XLS → PDF": "✅ 支持表格格式化输出",
            "XLS → XLSX": "✅ 格式升级转换",
            "XLSX → CSV": "✅ 支持所有工作表分别转换",
            "XLSX → PDF": "✅ 支持表格格式化输出",
            "XLSX → XLS": "⚠️ 建议使用XLSX格式"
        }
        
        for conversion, description in conversion_matrix.items():
            st.write(f"**{conversion}:** {description}")
    
    def _render_footer(self):
        """渲染页面底部"""
        st.divider()
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown(
                "<p style='text-align: center; color: #999; font-size: 0.8em;'>"
                "Excel批量转换工具 | 支持XLS/XLSX → CSV/PDF/XLSX"
                "</p>",
                unsafe_allow_html=True
            )
    
    def _cleanup_temp_files(self):
        """清理临时文件"""
        try:
            cleaned_count = FileManager.cleanup_old_temp_files()
            st.session_state.file_manager.cleanup_temp_files()
            
            # 重置状态
            st.session_state.uploaded_files = []
            st.session_state.conversion_config = None
            st.session_state.zip_path = None
            st.session_state.output_dir = None
            
            st.success(f"已清理 {cleaned_count} 个临时文件")
            st.rerun()
            
        except Exception as e:
            st.error(f"清理临时文件失败: {str(e)}")
    
    def _reset_application(self):
        """重置应用状态"""
        try:
            # 清理当前会话的文件
            st.session_state.file_manager.cleanup_temp_files()
            
            # 重置所有会话状态
            for key in ['uploaded_files', 'conversion_config', 'zip_path', 'output_dir']:
                if key in st.session_state:
                    del st.session_state[key]
            
            # 重新初始化
            self._init_session_state()
            
            st.success("应用已重置")
            st.rerun()
            
        except Exception as e:
            st.error(f"重置应用失败: {str(e)}")


def main():
    """主函数入口"""
    try:
        interface = MainInterface()
        interface.render()
    except Exception as e:
        st.error(f"应用启动失败: {str(e)}")
        st.stop()


if __name__ == "__main__":
    main() 