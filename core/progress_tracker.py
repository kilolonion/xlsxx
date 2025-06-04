# -*- coding: utf-8 -*-
"""
进度跟踪器模块
"""

import streamlit as st
import time
from typing import Dict, List, Optional, Callable, Any
from dataclasses import dataclass
from datetime import datetime, timedelta


@dataclass
class FileProgress:
    """单个文件进度信息"""
    file_name: str
    status: str = "等待中"  # 等待中、处理中、完成、失败
    progress: float = 0.0  # 0.0 - 1.0
    start_time: Optional[datetime] = None
    end_time: Optional[datetime] = None
    error_message: Optional[str] = None
    file_size: int = 0
    estimated_time: Optional[float] = None


class ProgressTracker:
    """进度跟踪器类"""
    
    def __init__(self, total_files: int, enable_time_estimation: bool = True):
        """
        初始化进度跟踪器
        
        Args:
            total_files: 总文件数
            enable_time_estimation: 是否启用时间估算
        """
        self.total_files = total_files
        self.enable_time_estimation = enable_time_estimation
        self.file_progress: List[FileProgress] = []
        self.overall_start_time = None
        self.processing_speeds: List[float] = []  # 每个文件的处理速度(MB/s)
        
        # Streamlit 组件
        self.progress_bar = None
        self.status_text = None
        self.details_container = None
        self.time_display = None
        
        # 统计信息
        self.completed_files = 0
        self.failed_files = 0
        self.total_size = 0
        self.processed_size = 0
    
    def initialize_files(self, file_info_list: List[Dict[str, Any]]) -> None:
        """
        初始化文件进度列表
        
        Args:
            file_info_list: 文件信息列表
        """
        self.file_progress = []
        self.total_size = 0
        
        for file_info in file_info_list:
            file_progress = FileProgress(
                file_name=file_info.get('name', '未知文件'),
                file_size=file_info.get('size', 0)
            )
            self.file_progress.append(file_progress)
            self.total_size += file_progress.file_size
    
    def setup_ui(self) -> None:
        """设置Streamlit UI组件"""
        # 总体进度条
        self.progress_bar = st.progress(0.0)
        self.status_text = st.empty()
        
        # 时间信息显示
        if self.enable_time_estimation:
            self.time_display = st.empty()
        
        # 详细进度显示
        with st.expander("详细进度", expanded=False):
            self.details_container = st.container()
    
    def start_processing(self) -> None:
        """开始处理"""
        self.overall_start_time = datetime.now()
        self.update_status("开始批量转换...")
    
    def start_file(self, file_index: int) -> None:
        """
        开始处理某个文件
        
        Args:
            file_index: 文件索引
        """
        if 0 <= file_index < len(self.file_progress):
            file_progress = self.file_progress[file_index]
            file_progress.status = "处理中"
            file_progress.start_time = datetime.now()
            file_progress.progress = 0.0
            
            self.update_status(f"正在处理: {file_progress.file_name}")
            self._update_ui()
    
    def update_file_progress(self, file_index: int, progress: float) -> None:
        """
        更新文件处理进度
        
        Args:
            file_index: 文件索引
            progress: 进度值 (0.0 - 1.0)
        """
        if 0 <= file_index < len(self.file_progress):
            self.file_progress[file_index].progress = max(0.0, min(1.0, progress))
            self._update_ui()
    
    def complete_file(self, file_index: int, success: bool = True, 
                     error_message: Optional[str] = None) -> None:
        """
        完成文件处理
        
        Args:
            file_index: 文件索引
            success: 是否成功
            error_message: 错误信息
        """
        if 0 <= file_index < len(self.file_progress):
            file_progress = self.file_progress[file_index]
            file_progress.end_time = datetime.now()
            file_progress.progress = 1.0
            
            if success:
                file_progress.status = "完成"
                self.completed_files += 1
                self.processed_size += file_progress.file_size
                
                # 计算处理速度
                if file_progress.start_time and file_progress.file_size > 0:
                    duration = (file_progress.end_time - file_progress.start_time).total_seconds()
                    if duration > 0:
                        speed = file_progress.file_size / (1024 * 1024) / duration  # MB/s
                        self.processing_speeds.append(speed)
            else:
                file_progress.status = "失败"
                file_progress.error_message = error_message
                self.failed_files += 1
            
            self._update_ui()
    
    def get_overall_progress(self) -> float:
        """获取总体进度"""
        if not self.file_progress:
            return 0.0
        
        total_progress = sum(fp.progress for fp in self.file_progress)
        return total_progress / len(self.file_progress)
    
    def get_estimated_time_remaining(self) -> Optional[str]:
        """获取预估剩余时间"""
        if not self.enable_time_estimation or not self.processing_speeds:
            return None
        
        # 计算平均处理速度
        avg_speed = sum(self.processing_speeds) / len(self.processing_speeds)
        
        # 计算剩余大小
        remaining_size = self.total_size - self.processed_size
        
        if avg_speed > 0 and remaining_size > 0:
            # 转换为MB
            remaining_mb = remaining_size / (1024 * 1024)
            estimated_seconds = remaining_mb / avg_speed
            
            # 格式化时间
            if estimated_seconds < 60:
                return f"{int(estimated_seconds)}秒"
            elif estimated_seconds < 3600:
                minutes = int(estimated_seconds / 60)
                seconds = int(estimated_seconds % 60)
                return f"{minutes}分{seconds}秒"
            else:
                hours = int(estimated_seconds / 3600)
                minutes = int((estimated_seconds % 3600) / 60)
                return f"{hours}时{minutes}分"
        
        return None
    
    def get_elapsed_time(self) -> str:
        """获取已消耗时间"""
        if not self.overall_start_time:
            return "0秒"
        
        elapsed = datetime.now() - self.overall_start_time
        total_seconds = int(elapsed.total_seconds())
        
        if total_seconds < 60:
            return f"{total_seconds}秒"
        elif total_seconds < 3600:
            minutes = total_seconds // 60
            seconds = total_seconds % 60
            return f"{minutes}分{seconds}秒"
        else:
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            return f"{hours}时{minutes}分"
    
    def update_status(self, message: str) -> None:
        """更新状态文本"""
        if self.status_text:
            self.status_text.text(message)
    
    def _update_ui(self) -> None:
        """更新UI显示"""
        # 更新总体进度条
        overall_progress = self.get_overall_progress()
        if self.progress_bar:
            self.progress_bar.progress(overall_progress)
        
        # 更新时间信息
        if self.time_display and self.enable_time_estimation:
            elapsed = self.get_elapsed_time()
            remaining = self.get_estimated_time_remaining()
            
            time_info = f"⏱️ 已耗时: {elapsed}"
            if remaining:
                time_info += f" | 预估剩余: {remaining}"
            
            self.time_display.info(time_info)
        
        # 更新详细进度
        if self.details_container:
            with self.details_container:
                self._render_detailed_progress()
    
    def _render_detailed_progress(self) -> None:
        """渲染详细进度信息"""
        # 状态统计 - 使用简单的文本显示
        waiting = self.total_files - self.completed_files - self.failed_files
        st.write("**进度统计:**")
        st.write(f"总文件: {self.total_files} | 已完成: {self.completed_files} | 失败: {self.failed_files} | 等待中: {waiting}")
        
        # 文件详细状态
        st.write("**文件处理状态:**")
        
        for i, fp in enumerate(self.file_progress):
            # 状态图标
            status_icons = {
                "等待中": "⏳",
                "处理中": "🔄", 
                "完成": "✅",
                "失败": "❌"
            }
            icon = status_icons.get(fp.status, "📄")
            
            # 简单的文本显示，避免嵌套列
            st.write(f"{icon} **{fp.file_name}** - {fp.status}")
            
            # 显示进度条
            st.progress(fp.progress)
            
            # 如果有错误信息，显示它
            if fp.status == "失败" and fp.error_message:
                st.error(f"错误: {fp.error_message}")
            
            st.write("")  # 添加一点间距
    
    def finish(self, success: bool = True) -> Dict[str, Any]:
        """
        完成处理并返回统计信息
        
        Args:
            success: 总体是否成功
            
        Returns:
            统计信息字典
        """
        end_time = datetime.now()
        total_time = None
        
        if self.overall_start_time:
            total_time = (end_time - self.overall_start_time).total_seconds()
        
        # 最终状态更新
        if success:
            self.update_status("✅ 批量转换完成！")
        else:
            self.update_status("❌ 批量转换失败")
        
        # 返回统计信息
        return {
            'total_files': self.total_files,
            'completed_files': self.completed_files,
            'failed_files': self.failed_files,
            'total_time': total_time,
            'total_size': self.total_size,
            'processed_size': self.processed_size,
            'average_speed': sum(self.processing_speeds) / len(self.processing_speeds) if self.processing_speeds else 0
        }
    
    def create_file_callback(self, file_index: int) -> Callable[[float], None]:
        """
        创建文件进度回调函数
        
        Args:
            file_index: 文件索引
            
        Returns:
            进度回调函数
        """
        def callback(progress: float):
            self.update_file_progress(file_index, progress)
        
        return callback 