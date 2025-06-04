#!/usr/bin/env python
"""
PDF模块集成器
提供对所有PDF处理辅助模块的集中访问，确保导入兼容性
"""

import os
import sys

# 添加当前目录到路径
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

# 色彩管理
def get_color_manager():
    """获取PDFColorManager实例"""
    try:
        from pdf_color_manager import PDFColorManager
        return PDFColorManager()
    except ImportError as e:
        print(f"警告: 无法导入PDFColorManager: {e}")
        return None

# CMYK颜色处理
def get_cmyk_helper():
    """获取CMYK处理助手"""
    try:
        import pdf_cmyk_helper
        return pdf_cmyk_helper
    except ImportError as e:
        print(f"警告: 无法导入pdf_cmyk_helper: {e}")
        return None

# 字体管理        
def get_font_manager():
    """获取PDFFontManager实例"""
    try:
        from pdf_font_manager import PDFFontManager
        return PDFFontManager()
    except ImportError as e:
        print(f"警告: 无法导入PDFFontManager: {e}")
        return None

# 处理CMYK颜色空间的Pixmap对象
def handle_pixmap_color(pixmap, adjustments=None):
    """
    处理CMYK颜色空间的Pixmap对象
    
    参数:
        pixmap: fitz.Pixmap对象
        adjustments: 颜色调整参数字典，默认为None
        
    返回:
        处理后的RGB颜色空间的Pixmap对象
    """
    cmyk_helper = get_cmyk_helper()
    if cmyk_helper:
        return cmyk_helper.handle_pixmap_color(pixmap, adjustments)
    else:
        # 简单的备用转换 - 如果cmyk_helper不可用
        import fitz
        if hasattr(pixmap, 'colorspace') and pixmap.colorspace and pixmap.colorspace.name in ("CMYK", "DeviceCMYK"):
            try:
                pix = fitz.Pixmap(fitz.csRGB, pixmap)
                return pix
            except Exception as e:
                print(f"颜色空间转换失败: {e}")
        return pixmap  # 返回原始pixmap，如果无法转换
