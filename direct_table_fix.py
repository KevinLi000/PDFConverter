"""
直接应用表格检测和提取修复
"""

import os
import sys
import types
import traceback

def apply_direct_table_fixes(converter):
    """
    直接应用表格检测和提取修复到转换器实例
    
    参数:
        converter: EnhancedPDFConverter实例
    """
    # 添加增强型表格检测方法
    def enhanced_detect_tables(self, page):
        """增强型表格检测方法"""
        try:
            # 首先尝试使用内置的find_tables方法
            try:
                import fitz
                tables = page.find_tables()
                if tables and hasattr(tables, 'tables') and len(tables.tables) > 0:
                    print(f"使用PyMuPDF内置方法检测到{len(tables.tables)}个表格")
                    return tables
            except (AttributeError, TypeError) as e:
                print(f"PyMuPDF的find_tables方法不可用 ({e})，使用增强检测")
            
            # 使用OpenCV检测表格
            tables = detect_tables_opencv(self, page)
            if tables and hasattr(tables, 'tables') and len(tables.tables) > 0:
                print(f"使用OpenCV方法检测到{len(tables.tables)}个表格")
                return tables
            
            # 使用布局分析检测表格
            tables = detect_tables_by_layout(self, page)
            if tables and hasattr(tables, 'tables') and len(tables.tables) > 0:
                print(f"使用布局分析方法检测到{len(tables.tables)}个表格")
                return tables
            
            # 如果所有方法都失败，创建一个空的表格集合
            class EmptyTableCollection:
                def __init__(self):
                    self.tables = []
            
            print("未检测到任何表格")
            return EmptyTableCollection()
            
        except Exception as e:
            print(f"表格检测错误: {e}")
            traceback.print_exc()
            
            # 返回空表格集合
            class EmptyTableCollection:
                def __init__(self):
                    self.tables = []
            
            return EmptyTableCollection()
    
    # OpenCV表格检测方法
    def detect_tables_opencv(self, page):
        """使用OpenCV检测表格"""
        try:
            import cv2
            import numpy as np
            from PIL import Image
            import fitz
            
            # 提高分辨率渲染页面为图像
            zoom = 3.0
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            
            # 转换为PIL图像
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # 转换为OpenCV格式
            img_np = np.array(img)
            gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
            
            # 使用更强的自适应阈值处理
            binary = cv2.adaptiveThreshold(
                gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2
            )
            
            # 应用形态学闭操作来连接线段
            kernel = np.ones((3,3), np.uint8)
            binary = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)
            
            # 寻找水平线 - 使用更灵活的参数
            horizontal = binary.copy()
            horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (30, 1))
            horizontal = cv2.morphologyEx(horizontal, cv2.MORPH_OPEN, horizontal_kernel)
            horizontal = cv2.dilate(horizontal, np.ones((1,5), np.uint8), iterations=1)
            
            # 寻找垂直线 - 使用更灵活的参数
            vertical = binary.copy()
            vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 30))
            vertical = cv2.morphologyEx(vertical, cv2.MORPH_OPEN, vertical_kernel)
            vertical = cv2.dilate(vertical, np.ones((5,1), np.uint8), iterations=1)
            
            # 合并水平和垂直线
            table_mask = cv2.bitwise_or(horizontal, vertical)
            
            # 应用连通组件分析来合并表格区域
            kernel = np.ones((5,5), np.uint8)
            table_mask = cv2.dilate(table_mask, kernel, iterations=3)
            
            # 寻找轮廓
            contours, _ = cv2.findContours(table_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            
            # 转换检测到的表格区域为PDF坐标
            tables = []
            page_width, page_height = page.rect.width, page.rect.height
            scale_x = page_width / pix.width
            scale_y = page_height / pix.height
            
            for contour in contours:
                x, y, w, h = cv2.boundingRect(contour)
                
                # 通过面积和纵横比过滤噪声区域
                area = w * h
                aspect_ratio = float(w) / h if h > 0 else 0
                
                # 过滤掉太小或形状不像表格的区域
                if area > 5000 and 0.1 < aspect_ratio < 10:
                    # 转换回PDF坐标
                    pdf_x0 = x * scale_x
                    pdf_y0 = y * scale_y
                    pdf_x1 = (x + w) * scale_x
                    pdf_y1 = (y + h) * scale_y
                    
                    # 创建表格对象
                    table = {
                        "bbox": (pdf_x0, pdf_y0, pdf_x1, pdf_y1),
                        "type": "table"
                    }
                    
                    tables.append(table)
            
            # 创建一个模拟的表格集合对象
            class TableCollection:
                def __init__(self, tables_list):
                    self.tables = tables_list
            
            return TableCollection(tables)
            
        except Exception as e:
            print(f"OpenCV表格检测错误: {e}")
            traceback.print_exc()
            return None
    
    # 基于布局的表格检测方法
    def detect_tables_by_layout(self, page):
        """使用文本块布局分析检测表格"""
        try:
            import fitz
            import numpy as np
            from collections import defaultdict
            
            # 获取页面文本块
            page_dict = page.get_text("dict")
            blocks = page_dict.get("blocks", [])
            
            # 收集可能是表格单元格的文本块
            potential_cells = []
            for block in blocks:
                if block["type"] == 0:  # 文本块
                    # 获取文本块的边界框和内容
                    x0, y0, x1, y1 = block["bbox"]
                    text = ""
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            text += span.get("text", "")
                    
                    # 只考虑包含文本的块
                    if text.strip():
                        potential_cells.append({
                            "bbox": (x0, y0, x1, y1),
                            "text": text
                        })
            
            # 如果找到的潜在单元格太少，可能没有表格
            if len(potential_cells) < 4:
                return None
            
            # 使用文本块的对齐方式检测表格
            # 按Y坐标对文本块分组，找到可能的表格行
            y_tolerance = page.rect.height * 0.01  # 容差为页面高度的1%
            y_groups = defaultdict(list)
            
            for cell in potential_cells:
                y_center = (cell["bbox"][1] + cell["bbox"][3]) / 2
                # 查找最接近的现有组
                found_group = False
                for group_y in sorted(y_groups.keys()):
                    if abs(group_y - y_center) < y_tolerance:
                        y_groups[group_y].append(cell)
                        found_group = True
                        break
                # 如果没有找到匹配的组，创建新组
                if not found_group:
                    y_groups[y_center].append(cell)
            
            # 如果至少有2行，每行至少有2个文本块，则可能是表格
            sorted_rows = sorted(y_groups.items(), key=lambda x: x[0])
            potential_table_rows = [row for _, row in sorted_rows if len(row) >= 2]
            
            if len(potential_table_rows) < 2:
                return None
            
            # 计算所有单元格的中心点，然后检查对齐情况
            all_centers_x = []
            for row in potential_table_rows:
                for cell in row:
                    center_x = (cell["bbox"][0] + cell["bbox"][2]) / 2
                    all_centers_x.append(center_x)
            
            # 使用聚类算法对X坐标分组
            x_tolerance = page.rect.width * 0.03  # 容差为页面宽度的3%
            x_groups = cluster_positions(all_centers_x, x_tolerance)
            
            # 如果X坐标分组少于2个，可能不是表格
            if len(x_groups) < 2:
                return None
            
            # 确定表格区域
            min_x = min([cell["bbox"][0] for row in potential_table_rows for cell in row])
            max_x = max([cell["bbox"][2] for row in potential_table_rows for cell in row])
            min_y = min([cell["bbox"][1] for row in potential_table_rows for cell in row])
            max_y = max([cell["bbox"][3] for row in potential_table_rows for cell in row])
            
            # 略微扩大表格边界
            padding = min(page.rect.width, page.rect.height) * 0.01
            table = {
                "bbox": (max(0, min_x - padding), 
                         max(0, min_y - padding), 
                         min(page.rect.width, max_x + padding), 
                         min(page.rect.height, max_y + padding)),
                "type": "table",
                "rows": len(potential_table_rows),
                "cols": len(x_groups)
            }
            
            # 创建一个模拟的表格集合对象
            class TableCollection:
                def __init__(self, tables_list):
                    self.tables = tables_list
            
            return TableCollection([table])
            
        except Exception as e:
            print(f"布局分析表格检测错误: {e}")
            traceback.print_exc()
            return None
    
    # 添加提取表格方法
    def extract_tables(self, pdf_document, page_num):
        """从PDF页面提取表格"""
        try:
            page = pdf_document[page_num]
            
            # 使用增强的表格检测
            if hasattr(self, 'detect_tables'):
                # 使用增强的detect_tables方法
                tables_obj = self.detect_tables(page)
                if tables_obj and hasattr(tables_obj, 'tables'):
                    return tables_obj.tables
                else:
                    return []
            
            # 备用方法：尝试使用find_tables (如果可用)
            try:
                tables = page.find_tables()
                if tables and len(tables.tables) > 0:
                    return tables.tables
            except (AttributeError, TypeError) as e:
                print(f"表格检测警告: {e}")
            
            return []
            
        except Exception as e:
            print(f"表格提取错误: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    # 辅助函数：聚类位置值
    def cluster_positions(positions, tolerance=5):
        """将接近的位置值聚类"""
        if not positions:
            return []
        
        # 排序位置
        sorted_pos = sorted(positions)
        
        # 初始化聚类
        clusters = [[sorted_pos[0]]]
        
        # 聚类过程
        for pos in sorted_pos[1:]:
            # 检查是否与上一个聚类接近
            if pos - clusters[-1][-1] <= tolerance:
                # 添加到现有聚类
                clusters[-1].append(pos)
            else:
                # 创建新聚类
                clusters.append([pos])
        
        # 对每个聚类取平均值
        cluster_centers = [sum(cluster) / len(cluster) for cluster in clusters]
        
        return cluster_centers
    
    # 绑定方法到转换器
    converter.detect_tables = types.MethodType(enhanced_detect_tables, converter)
    converter.detect_tables_opencv = types.MethodType(detect_tables_opencv, converter)
    converter.detect_tables_by_layout = types.MethodType(detect_tables_by_layout, converter)
    converter._extract_tables = types.MethodType(extract_tables, converter)
    
    # 将辅助函数添加到模块全局变量中
    globals()['cluster_positions'] = cluster_positions
    
    print("已应用直接表格检测和提取修复")
    return True

def main():
    """主函数"""
    try:
        # 导入增强型PDF转换器
        from enhanced_pdf_converter import EnhancedPDFConverter
        
        # 创建转换器实例
        converter = EnhancedPDFConverter()
        
        # 应用表格修复
        apply_direct_table_fixes(converter)
        
        # 检查是否成功添加了表格检测和提取方法
        has_detect_tables = hasattr(converter, 'detect_tables')
        has_detect_tables_opencv = hasattr(converter, 'detect_tables_opencv')
        has_detect_tables_by_layout = hasattr(converter, 'detect_tables_by_layout')
        has_extract_tables = hasattr(converter, '_extract_tables')
        
        # 输出结果
        print("===== 表格检测和提取修复测试结果 =====")
        print(f"检测到detect_tables方法: {'是' if has_detect_tables else '否'}")
        print(f"检测到detect_tables_opencv方法: {'是' if has_detect_tables_opencv else '否'}")
        print(f"检测到detect_tables_by_layout方法: {'是' if has_detect_tables_by_layout else '否'}")
        print(f"检测到_extract_tables方法: {'是' if has_extract_tables else '否'}")
        print("=====================================")
        
        return True
        
    except Exception as e:
        print(f"应用表格修复失败: {e}")
        traceback.print_exc()
        return False

if __name__ == "__main__":
    result = main()
    exit(0 if result else 1)
