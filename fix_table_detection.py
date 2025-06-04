"""
增强PDF转换器表格检测补丁
修复'Page' object has no attribute 'find_tables'错误
"""

import os
import sys
import traceback

# 检查必要的依赖
try:
    import fitz  # PyMuPDF
    import cv2
    import numpy as np
    from PIL import Image
except ImportError as e:
    print(f"错误: 无法导入所需库: {e}")
    print("请安装必要的库: pip install PyMuPDF opencv-python numpy pillow")
    sys.exit(1)

def apply_table_detection_patch():
    """
    应用表格检测补丁，修复find_tables错误
    """
    # 检查PDF转换器
    from enhanced_pdf_converter import EnhancedPDFConverter
    
    # 给转换器添加增强的表格检测方法
    def detect_tables(self, page):
        """
        增强的表格检测方法，支持不同版本的PyMuPDF
        
        参数:
            page: fitz.Page对象
            
        返回:
            表格区域列表
        """
        try:
            # 首先尝试使用内置的find_tables方法
            try:
                tables = page.find_tables()
                if tables and hasattr(tables, 'tables'):
                    return tables
                return []
            except (AttributeError, TypeError) as e:
                # 如果find_tables不可用，使用备用方法
                print(f"PyMuPDF的find_tables方法不可用 ({e})，使用备用表格检测")
                
                # 使用OpenCV检测表格
                tables = extract_tables_opencv(page, dpi=self.dpi if hasattr(self, 'dpi') else 300)
                return tables
        except Exception as e:
            print(f"表格检测错误: {e}")
            traceback.print_exc()
            return []
    
    # 添加OpenCV表格检测
    def extract_tables_opencv(page, dpi=300):
        """
        使用OpenCV检测表格
        
        参数:
            page: fitz.Page对象
            dpi: 渲染DPI
            
        返回:
            表格区域列表
        """
        try:
            # 提高分辨率渲染页面为图像
            zoom = dpi / 72  # 计算缩放比例
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            
            # 转换为PIL图像
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # 转换为OpenCV格式
            img_np = np.array(img)
            gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
            
            # 自适应阈值处理
            binary = cv2.adaptiveThreshold(
                gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 15, 2
            )
            
            # 寻找水平线
            horizontal = binary.copy()
            horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
            horizontal = cv2.morphologyEx(horizontal, cv2.MORPH_OPEN, horizontal_kernel)
            
            # 寻找垂直线
            vertical = binary.copy()
            vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
            vertical = cv2.morphologyEx(vertical, cv2.MORPH_OPEN, vertical_kernel)
            
            # 合并水平和垂直线
            table_mask = cv2.bitwise_or(horizontal, vertical)
            
            # 寻找轮廓
            contours, _ = cv2.findContours(table_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            
            # 转换检测到的表格区域为PDF坐标
            tables = []
            page_width, page_height = page.rect.width, page.rect.height
            scale_x = page_width / pix.width
            scale_y = page_height / pix.height
            
            for contour in contours:
                x, y, w, h = cv2.boundingRect(contour)
                # 过滤太小的区域
                if w > 50 and h > 50:
                    # 转换回PDF坐标
                    pdf_x0 = x * scale_x
                    pdf_y0 = y * scale_y
                    pdf_x1 = (x + w) * scale_x
                    pdf_y1 = (y + h) * scale_y
                    
                    # 创建类表格对象结构
                    table = {
                        "bbox": (pdf_x0, pdf_y0, pdf_x1, pdf_y1),
                        "type": "table"
                    }
                    
                    tables.append(table)
            
            # 创建一个模拟的表格集合对象，与PyMuPDF的find_tables()返回结构兼容
            class TableCollection:
                def __init__(self, tables_list):
                    self.tables = tables_list
            
            return TableCollection(tables)
        except Exception as e:
            print(f"OpenCV表格检测错误: {e}")
            traceback.print_exc()
            return []
    
    # 绑定方法到转换器类
    import types
    EnhancedPDFConverter.detect_tables = detect_tables
    
    print("已应用表格检测补丁，解决了'Page' object has no attribute 'find_tables'错误")
    return True

# 如果作为独立脚本运行，则应用补丁
if __name__ == "__main__":
    apply_table_detection_patch()
