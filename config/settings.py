# -*- coding: utf-8 -*-
"""
配置文件
定义应用程序的所有常量和配置参数
"""

import os

# 文件处理配置
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
SUPPORTED_INPUT_FORMATS = ['xls', 'xlsx']
SUPPORTED_OUTPUT_FORMATS = ['xlsx', 'csv', 'pdf']

# 临时文件配置
TEMP_DIR_CLEANUP_HOURS = 24
TEMP_DIR_BASE = "temp"

# 批量处理配置
MAX_FILES_PER_BATCH = 10
MAX_CONCURRENT_CONVERSIONS = 3

# PDF配置
PDF_PAGE_SIZES = {
    'A4': (210, 297),  # mm
    'A3': (297, 420),  # mm
    'Letter': (216, 279)  # mm
}

PDF_ORIENTATIONS = ['纵向', '横向']

# CSV配置
CSV_ENCODINGS = ['UTF-8', 'GBK', 'ISO-8859-1']
CSV_SEPARATORS = {
    '逗号 (,)': ',',
    '分号 (;)': ';',
    '制表符': '\t',
    '竖线 (|)': '|'
}

# 应用配置
APP_TITLE = "XLS批量转换工具"
APP_ICON = "📊"
APP_DESCRIPTION = "支持XLS、XLSX文件批量转换为CSV、PDF、XLSX格式"

# XLSX配置
XLSX_OPTIONS = {
    '保留原始格式': True,
    '保留公式': False,
    '工作表保护': False,
    '内存优化模式': True
}

XLSX_COMPRESSION_LEVELS = {
    '无压缩': 0,
    '快速压缩': 1,
    '标准压缩': 6,
    '最大压缩': 9
}

# 界面配置
UPLOAD_HELP_TEXT = "支持拖拽上传多个文件，单个文件不超过50MB"
MAX_PREVIEW_ROWS = 10 