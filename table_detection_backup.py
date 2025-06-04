"""
增强表格检测功能，可作为备用方法使用
"""

import fitz
import numpy as np
from PIL import Image
import os
import tempfile

def extract_tables_opencv(page, dpi=300):
    """
    使用OpenCV进行表格检测的备用方法
    
    参数:
        page: PyMuPDF页面对象
        dpi: 图像分辨率
        
    返回:
        检测到的表格列表
    """
    try:
        import cv2
    except ImportError:
        print("需要安装OpenCV库: pip install opencv-python")
        return []
    
    # 提高分辨率渲染页面为图像
    zoom = dpi / 72  # 转换DPI为缩放因子
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    
    # 转换为PIL图像
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    
    # 转换为OpenCV格式
    img_np = np.array(img)
    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    
    # 使用自适应阈值处理
    binary = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 15, 2
    )
    
    # 查找水平和垂直线
    horizontal = binary.copy()
    vertical = binary.copy()
    
    # 处理水平线
    horizontalStructure = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
    horizontal = cv2.morphologyEx(horizontal, cv2.MORPH_OPEN, horizontalStructure)
    
    # 处理垂直线
    verticalStructure = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
    vertical = cv2.morphologyEx(vertical, cv2.MORPH_OPEN, verticalStructure)
    
    # 合并水平线和垂直线
    table_mask = cv2.bitwise_or(horizontal, vertical)
    
    # 查找轮廓 - 这些是潜在的表格
    contours, _ = cv2.findContours(table_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # 提取表格区域
    tables = []
    page_width, page_height = page.rect.width, page.rect.height
    scale_x, scale_y = page_width / pix.width, page_height / pix.height
    
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        # 过滤掉太小的区域
        if w > 100 and h > 100:
            # 转换回PDF坐标系
            pdf_x0 = x * scale_x
            pdf_y0 = y * scale_y
            pdf_x1 = (x + w) * scale_x
            pdf_y1 = (y + h) * scale_y
            
            # 创建表格结构
            table = {
                "bbox": (pdf_x0, pdf_y0, pdf_x1, pdf_y1),
                "type": "table"
            }
            
            # 提取表格结构（行和列）
            rows, cols = extract_table_structure(page, (pdf_x0, pdf_y0, pdf_x1, pdf_y1))
            if rows and cols:
                table["rows"] = rows
                table["cols"] = cols
                tables.append(table)
    
    return tables

def extract_table_structure(page, table_rect):
    """
    提取表格的行和列结构
    
    参数:
        page: PDF页面对象
        table_rect: 表格区域 (x0, y0, x1, y1)
        
    返回:
        rows, cols: 行和列的位置列表
    """
    try:
        # 从表格区域提取文本块
        clip_rect = fitz.Rect(table_rect)
        blocks = page.get_text("dict", clip=clip_rect)["blocks"]
        
        # 收集所有文本行的位置，用于确定行
        all_lines = []
        for block in blocks:
            if block["type"] == 0:  # 文本块
                for line in block.get("lines", []):
                    y0, y1 = line["bbox"][1], line["bbox"][3]
                    all_lines.append((y0, y1))
        
        # 对行位置进行分组
        rows = cluster_positions([line[0] for line in all_lines], tolerance=5)
        
        # 提取所有span的左侧位置，用于确定列
        all_spans_x0 = []
        for block in blocks:
            if block["type"] == 0:
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        all_spans_x0.append(span["bbox"][0])
        
        # 对列位置进行分组
        cols = cluster_positions(all_spans_x0, tolerance=10)
        
        # 确保第一列和最后一列包括表格边界
        if cols and cols[0] > table_rect[0] + 5:
            cols.insert(0, table_rect[0])
        if cols and cols[-1] < table_rect[2] - 5:
            cols.append(table_rect[2])
        
        # 确保第一行和最后一行包括表格边界
        if rows and rows[0] > table_rect[1] + 5:
            rows.insert(0, table_rect[1])
        if rows and rows[-1] < table_rect[3] - 5:
            rows.append(table_rect[3])
        
        return rows, cols
    except Exception as e:
        print(f"提取表格结构错误: {e}")
        return [], []

def cluster_positions(positions, tolerance=5):
    """
    对位置值进行聚类，用于确定表格的行和列
    
    参数:
        positions: 位置值列表
        tolerance: 聚类容差
        
    返回:
        list: 聚类后的位置值
    """
    if not positions:
        return []
        
    # 排序位置
    sorted_positions = sorted(positions)
    
    # 初始化聚类
    clusters = [[sorted_positions[0]]]
    
    # 聚类过程
    for pos in sorted_positions[1:]:
        last_cluster = clusters[-1]
        last_cluster_avg = sum(last_cluster) / len(last_cluster)
        
        if abs(pos - last_cluster_avg) <= tolerance:
            # 添加到最后一个聚类
            last_cluster.append(pos)
        else:
            # 创建新聚类
            clusters.append([pos])
    
    # 计算每个聚类的平均值
    return [sum(cluster) / len(cluster) for cluster in clusters]
