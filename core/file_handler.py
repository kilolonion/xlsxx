# -*- coding: utf-8 -*-
"""
文件处理器模块
负责文件的保存、管理和清理
"""

import os
import shutil
import zipfile
import tempfile
from typing import List, Optional, Tuple
from pathlib import Path
import streamlit as st

from .utils import (
    generate_session_id, 
    get_timestamp, 
    sanitize_filename, 
    ensure_directory_exists,
    clean_old_files
)
from config.settings import TEMP_DIR_BASE, TEMP_DIR_CLEANUP_HOURS


class FileManager:
    """文件管理器类"""
    
    def __init__(self, session_id: Optional[str] = None):
        """
        初始化文件管理器
        
        Args:
            session_id: 会话ID，如果为None则自动生成
        """
        self.session_id = session_id or generate_session_id()
        self.temp_dir = self._get_session_temp_dir()
        self._ensure_temp_dir_exists()
    
    def _get_session_temp_dir(self) -> str:
        """获取当前会话的临时目录路径"""
        return os.path.join(TEMP_DIR_BASE, self.session_id)
    
    def _ensure_temp_dir_exists(self) -> bool:
        """确保临时目录存在"""
        return ensure_directory_exists(self.temp_dir)
    
    def save_uploaded_file(self, uploaded_file, subfolder: str = "uploads") -> Optional[str]:
        """
        保存上传的文件到临时目录
        
        Args:
            uploaded_file: Streamlit上传的文件对象
            subfolder: 子文件夹名称
            
        Returns:
            保存的文件路径，失败时返回None
        """
        if uploaded_file is None:
            return None
        
        try:
            # 创建子目录
            upload_dir = os.path.join(self.temp_dir, subfolder)
            ensure_directory_exists(upload_dir)
            
            # 清理文件名
            safe_filename = sanitize_filename(uploaded_file.name)
            
            # 添加时间戳避免重名
            name_parts = os.path.splitext(safe_filename)
            timestamped_name = f"{name_parts[0]}_{get_timestamp()}{name_parts[1]}"
            
            # 完整文件路径
            file_path = os.path.join(upload_dir, timestamped_name)
            
            # 保存文件
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            return file_path
            
        except Exception as e:
            st.error(f"保存文件失败: {str(e)}")
            return None
    
    def create_temp_directory(self, subfolder: str = "") -> str:
        """
        创建临时子目录
        
        Args:
            subfolder: 子文件夹名称
            
        Returns:
            创建的目录路径
        """
        if subfolder:
            dir_path = os.path.join(self.temp_dir, subfolder)
        else:
            dir_path = self.temp_dir
        
        ensure_directory_exists(dir_path)
        return dir_path
    
    def cleanup_temp_files(self) -> bool:
        """
        清理当前会话的临时文件
        
        Returns:
            清理是否成功
        """
        try:
            if os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
                return True
            return True
        except Exception as e:
            st.error(f"清理临时文件失败: {str(e)}")
            return False
    
    def zip_converted_files(self, file_list: List[str], zip_name: str = None) -> Optional[str]:
        """
        将转换后的文件打包为ZIP
        
        Args:
            file_list: 要打包的文件路径列表
            zip_name: ZIP文件名，如果为None则自动生成
            
        Returns:
            ZIP文件路径，失败时返回None
        """
        if not file_list:
            return None
        
        try:
            # 生成ZIP文件名
            if zip_name is None:
                zip_name = f"converted_files_{get_timestamp()}.zip"
            
            # 确保ZIP文件名安全
            safe_zip_name = sanitize_filename(zip_name)
            if not safe_zip_name.endswith('.zip'):
                safe_zip_name += '.zip'
            
            # ZIP文件路径
            zip_path = os.path.join(self.temp_dir, safe_zip_name)
            
            # 创建ZIP文件
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file_path in file_list:
                    if os.path.exists(file_path):
                        # 使用文件名作为ZIP内的路径
                        arcname = os.path.basename(file_path)
                        zipf.write(file_path, arcname)
            
            return zip_path if os.path.exists(zip_path) else None
            
        except Exception as e:
            st.error(f"创建ZIP文件失败: {str(e)}")
            return None
    
    def get_file_info(self, file_path: str) -> dict:
        """
        获取文件信息
        
        Args:
            file_path: 文件路径
            
        Returns:
            包含文件信息的字典
        """
        info = {
            'name': '',
            'size': 0,
            'size_formatted': '0B',
            'extension': '',
            'exists': False
        }
        
        try:
            if os.path.exists(file_path):
                info['exists'] = True
                info['name'] = os.path.basename(file_path)
                info['size'] = os.path.getsize(file_path)
                info['size_formatted'] = self._format_file_size(info['size'])
                info['extension'] = os.path.splitext(file_path)[1].lower()
        except Exception:
            pass
        
        return info
    
    def _format_file_size(self, size_bytes: int) -> str:
        """格式化文件大小"""
        if size_bytes == 0:
            return "0B"
        
        size_names = ["B", "KB", "MB", "GB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
        
        return f"{size_bytes:.1f}{size_names[i]}"
    
    def get_individual_file_info(self, file_path: str) -> Optional[dict]:
        """
        获取单个文件的详细信息
        
        Args:
            file_path: 文件路径
            
        Returns:
            文件信息字典或None
        """
        if not os.path.exists(file_path):
            return None
        
        try:
            file_info = self.get_file_info(file_path)
            # 添加下载所需的额外信息
            file_info['download_name'] = os.path.basename(file_path)
            file_info['mime_type'] = self._get_mime_type(file_path)
            
            return file_info
            
        except Exception as e:
            st.error(f"获取文件信息失败: {str(e)}")
            return None
    
    def generate_individual_download_data(self, file_path: str) -> Optional[bytes]:
        """
        生成单个文件的下载数据
        
        Args:
            file_path: 文件路径
            
        Returns:
            文件二进制数据或None
        """
        try:
            if not os.path.exists(file_path):
                return None
                
            with open(file_path, 'rb') as f:
                return f.read()
                
        except Exception as e:
            st.error(f"读取文件失败: {str(e)}")
            return None
    
    def list_converted_files(self, output_dir: str) -> List[dict]:
        """
        列出转换完成的文件
        
        Args:
            output_dir: 输出目录路径
            
        Returns:
            文件信息列表
        """
        files_info = []
        
        try:
            if not os.path.exists(output_dir):
                return files_info
                
            for filename in os.listdir(output_dir):
                file_path = os.path.join(output_dir, filename)
                if os.path.isfile(file_path):
                    file_info = self.get_individual_file_info(file_path)
                    if file_info:
                        file_info['full_path'] = file_path
                        files_info.append(file_info)
                        
        except Exception as e:
            st.error(f"列出文件失败: {str(e)}")
            
        return files_info
    
    def _get_mime_type(self, file_path: str) -> str:
        """获取文件的MIME类型"""
        extension = os.path.splitext(file_path)[1].lower()
        
        mime_types = {
            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            '.csv': 'text/csv',
            '.pdf': 'application/pdf',
            '.xls': 'application/vnd.ms-excel'
        }
        
        return mime_types.get(extension, 'application/octet-stream')

    @staticmethod
    def cleanup_old_temp_files() -> int:
        """
        清理所有过期的临时文件
        
        Returns:
            清理的文件数量
        """
        if not os.path.exists(TEMP_DIR_BASE):
            return 0
        
        return clean_old_files(TEMP_DIR_BASE, TEMP_DIR_CLEANUP_HOURS) 