# -*- coding: utf-8 -*-
"""
工具函数模块
提供通用的辅助函数
"""

import os
import uuid
import datetime
import hashlib
from typing import List, Optional
from pathlib import Path

def generate_session_id() -> str:
    """生成唯一的会话ID"""
    return str(uuid.uuid4())

def get_timestamp() -> str:
    """获取当前时间戳字符串"""
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

def format_file_size(size_bytes: int) -> str:
    """格式化文件大小显示"""
    if size_bytes == 0:
        return "0B"
    
    size_names = ["B", "KB", "MB", "GB"]
    i = 0
    while size_bytes >= 1024 and i < len(size_names) - 1:
        size_bytes /= 1024.0
        i += 1
    
    return f"{size_bytes:.1f}{size_names[i]}"

def validate_file_extension(filename: str, allowed_extensions: List[str]) -> bool:
    """验证文件扩展名"""
    if not filename:
        return False
    
    file_ext = Path(filename).suffix.lower().lstrip('.')
    return file_ext in [ext.lower() for ext in allowed_extensions]

def sanitize_filename(filename: str) -> str:
    """清理文件名，移除非法字符"""
    # 移除或替换非法字符
    illegal_chars = '<>:"/\\|?*'
    for char in illegal_chars:
        filename = filename.replace(char, '_')
    
    # 移除多余的空格和点
    filename = filename.strip(' .')
    
    return filename

def calculate_file_hash(file_path: str) -> str:
    """计算文件的MD5哈希值"""
    hash_md5 = hashlib.md5()
    try:
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    except Exception:
        return ""

def ensure_directory_exists(directory_path: str) -> bool:
    """确保目录存在，不存在则创建"""
    try:
        Path(directory_path).mkdir(parents=True, exist_ok=True)
        return True
    except Exception:
        return False

def clean_old_files(directory: str, hours_old: int = 24) -> int:
    """清理指定时间之前的旧文件"""
    if not os.path.exists(directory):
        return 0
    
    cutoff_time = datetime.datetime.now() - datetime.timedelta(hours=hours_old)
    cleaned_count = 0
    
    try:
        for root, dirs, files in os.walk(directory):
            for file in files:
                file_path = os.path.join(root, file)
                file_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path))
                
                if file_time < cutoff_time:
                    try:
                        os.remove(file_path)
                        cleaned_count += 1
                    except Exception:
                        continue
                        
            # 清理空目录
            for dir_name in dirs:
                dir_path = os.path.join(root, dir_name)
                try:
                    if not os.listdir(dir_path):
                        os.rmdir(dir_path)
                except Exception:
                    continue
                    
    except Exception:
        pass
    
    return cleaned_count 