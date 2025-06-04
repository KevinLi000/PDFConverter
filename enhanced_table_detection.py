"""
增强型表格检测补丁 - 专门针对表格识别不足的问题
"""

import os
import sys
import traceback
import types

def apply_enhanced_table_detection_patch(converter):
    """应用增强型表格检测补丁到转换器"""
    
    def enhanced_detect_tables(self, page):
        """
        增强型表格检测方法，专注于提高表格识别率
        
        参数:
            page: fitz.Page对象
            
        返回:
            表格区域列表
        """
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
            
            # 使用多种方法检测表格
            tables = []
            
            # 方法1: 使用OpenCV检测表格边框
            try:
                cv_tables = detect_tables_opencv(self, page)
                if cv_tables and hasattr(cv_tables, 'tables') and len(cv_tables.tables) > 0:
                    print(f"使用OpenCV方法检测到{len(cv_tables.tables)}个表格")
                    tables = cv_tables
            except Exception as e:
                print(f"OpenCV表格检测错误: {e}")
            
            # 方法2: 使用布局分析检测表格
            if not tables or len(tables.tables) == 0:
                try:
                    layout_tables = detect_tables_by_layout(self, page)
                    if layout_tables and hasattr(layout_tables, 'tables') and len(layout_tables.tables) > 0:
                        print(f"使用布局分析方法检测到{len(layout_tables.tables)}个表格")
                        tables = layout_tables
                except Exception as e:
                    print(f"布局分析表格检测错误: {e}")
            
            # 方法3: 使用规则网格检测表格
            if not tables or len(tables.tables) == 0:
                try:
                    grid_tables = detect_tables_by_grid(self, page)
                    if grid_tables and hasattr(grid_tables, 'tables') and len(grid_tables.tables) > 0:
                        print(f"使用规则网格方法检测到{len(grid_tables.tables)}个表格")
                        tables = grid_tables
                except Exception as e:
                    print(f"规则网格表格检测错误: {e}")
            
            # 方法4: 使用文本对齐检测表格
            if not tables or len(tables.tables) == 0:
                try:
                    text_align_tables = detect_tables_by_text_alignment(self, page)
                    if text_align_tables and hasattr(text_align_tables, 'tables') and len(text_align_tables.tables) > 0:
                        print(f"使用文本对齐方法检测到{len(text_align_tables.tables)}个表格")
                        tables = text_align_tables
                except Exception as e:
                    print(f"文本对齐表格检测错误: {e}")
            
            # 如果找到表格，返回结果
            if tables and hasattr(tables, 'tables') and len(tables.tables) > 0:
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

    def detect_tables_opencv(self, page):
        """
        使用OpenCV检测表格
        
        参数:
            page: fitz.Page对象
            
        返回:
            表格区域列表
        """
        try:
            import cv2
            import numpy as np
            from PIL import Image
            import fitz
            
            # 提高分辨率渲染页面为图像
            zoom = 3.0  # 增加放大因子，提高检测精度
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
                    
                    # 添加表格结构分析
                    analyze_table_structure(self, page, table)
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

    def detect_tables_by_layout(self, page):
        """
        使用文本块布局分析检测表格
        
        参数:
            page: fitz.Page对象
            
        返回:
            表格区域列表
        """
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
            # 1. 按Y坐标对文本块分组，找到可能的表格行
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
            
            # 2. 如果至少有2行，每行至少有2个文本块，则可能是表格
            sorted_rows = sorted(y_groups.items(), key=lambda x: x[0])
            potential_table_rows = [row for _, row in sorted_rows if len(row) >= 2]
            
            if len(potential_table_rows) < 2:
                return None
            
            # 3. 判断文本块是否形成网格结构
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
            
            # 4. 确定表格区域
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

    def detect_tables_by_grid(self, page):
        """
        使用规则网格检测表格
        
        参数:
            page: fitz.Page对象
            
        返回:
            表格区域列表
        """
        try:
            import fitz
            import numpy as np
            from collections import defaultdict
            
            # 获取页面文本
            page_dict = page.get_text("dict")
            blocks = page_dict.get("blocks", [])
            
            # 收集所有文本行
            all_lines = []
            for block in blocks:
                if block["type"] == 0:  # 文本块
                    for line in block.get("lines", []):
                        x0, y0, x1, y1 = line["bbox"]
                        text = ""
                        for span in line.get("spans", []):
                            text += span.get("text", "")
                        if text.strip():
                            all_lines.append({
                                "bbox": (x0, y0, x1, y1),
                                "text": text
                            })
            
            # 如果找到的文本行太少，可能没有表格
            if len(all_lines) < 4:
                return None
            
            # 1. 计算行间距
            sorted_lines = sorted(all_lines, key=lambda x: x["bbox"][1])
            line_gaps = []
            for i in range(1, len(sorted_lines)):
                current_top = sorted_lines[i]["bbox"][1]
                prev_bottom = sorted_lines[i-1]["bbox"][3]
                gap = current_top - prev_bottom
                if gap > 0:
                    line_gaps.append(gap)
            
            if not line_gaps:
                return None
            
            # 计算行间距的中位数
            median_gap = np.median(line_gaps)
            
            # 2. 查找规则间隔的行分组
            y_tolerance = median_gap * 0.5  # 容差为中位数间距的一半
            consistent_rows = []
            current_row = [sorted_lines[0]]
            
            for i in range(1, len(sorted_lines)):
                current_line = sorted_lines[i]
                prev_line = sorted_lines[i-1]
                
                gap = current_line["bbox"][1] - prev_line["bbox"][3]
                
                # 如果间距在容差范围内，认为是规则的行间距
                if abs(gap - median_gap) <= y_tolerance:
                    # 新行
                    consistent_rows.append(current_row)
                    current_row = [current_line]
                else:
                    # 同一行
                    current_row.append(current_line)
            
            # 添加最后一行
            if current_row:
                consistent_rows.append(current_row)
            
            # 3. 查找具有规则行数的表格区域
            # 如果至少有3行规则间隔的行，可能是表格
            if len(consistent_rows) < 3:
                return None
            
            # 4. 计算这些行的覆盖区域
            min_x = min([line["bbox"][0] for row in consistent_rows for line in row])
            max_x = max([line["bbox"][2] for row in consistent_rows for line in row])
            min_y = min([line["bbox"][1] for row in consistent_rows for line in row])
            max_y = max([line["bbox"][3] for row in consistent_rows for line in row])
            
            # 略微扩大表格边界
            padding = min(page.rect.width, page.rect.height) * 0.01
            table = {
                "bbox": (max(0, min_x - padding), 
                         max(0, min_y - padding), 
                         min(page.rect.width, max_x + padding), 
                         min(page.rect.height, max_y + padding)),
                "type": "table",
                "rows": len(consistent_rows),
                "cols": 0  # 列数稍后分析
            }
            
            # 创建一个模拟的表格集合对象
            class TableCollection:
                def __init__(self, tables_list):
                    self.tables = tables_list
            
            return TableCollection([table])
            
        except Exception as e:
            print(f"规则网格表格检测错误: {e}")
            traceback.print_exc()
            return None

    def detect_tables_by_text_alignment(self, page):
        """
        使用文本对齐特征检测表格
        
        参数:
            page: fitz.Page对象
            
        返回:
            表格区域列表
        """
        try:
            import fitz
            import numpy as np
            from collections import defaultdict
            
            # 获取页面文本
            page_dict = page.get_text("dict")
            blocks = page_dict.get("blocks", [])
            
            # 收集所有文本行
            all_lines = []
            for block in blocks:
                if block["type"] == 0:  # 文本块
                    for line in block.get("lines", []):
                        x0, y0, x1, y1 = line["bbox"]
                        text = ""
                        for span in line.get("spans", []):
                            text += span.get("text", "")
                        if text.strip():
                            all_lines.append({
                                "bbox": (x0, y0, x1, y1),
                                "text": text,
                                "start_x": x0,
                                "center_x": (x0 + x1) / 2
                            })
            
            # 如果找到的文本行太少，可能没有表格
            if len(all_lines) < 4:
                return None
            
            # 1. 检查垂直对齐的文本
            x_tolerance = page.rect.width * 0.02  # 容差为页面宽度的2%
            
            # 按起始X坐标对文本行分组
            x_start_groups = defaultdict(list)
            for line in all_lines:
                # 查找最接近的现有组
                found_group = False
                for group_x in sorted(x_start_groups.keys()):
                    if abs(group_x - line["start_x"]) < x_tolerance:
                        x_start_groups[group_x].append(line)
                        found_group = True
                        break
                # 如果没有找到匹配的组，创建新组
                if not found_group:
                    x_start_groups[line["start_x"]].append(line)
            
            # 2. 查找具有多个垂直对齐文本行的组
            aligned_groups = [group for group in x_start_groups.values() if len(group) >= 3]
            
            # 如果没有足够的垂直对齐组，可能没有表格
            if len(aligned_groups) < 2:
                return None
            
            # 3. 找出这些对齐组的覆盖区域
            all_aligned_lines = [line for group in aligned_groups for line in group]
            min_x = min([line["bbox"][0] for line in all_aligned_lines])
            max_x = max([line["bbox"][2] for line in all_aligned_lines])
            min_y = min([line["bbox"][1] for line in all_aligned_lines])
            max_y = max([line["bbox"][3] for line in all_aligned_lines])
            
            # 略微扩大表格边界
            padding = min(page.rect.width, page.rect.height) * 0.01
            table = {
                "bbox": (max(0, min_x - padding), 
                         max(0, min_y - padding), 
                         min(page.rect.width, max_x + padding), 
                         min(page.rect.height, max_y + padding)),
                "type": "table",
                "cols": len(aligned_groups)
            }
            
            # 创建一个模拟的表格集合对象
            class TableCollection:
                def __init__(self, tables_list):
                    self.tables = tables_list
            
            return TableCollection([table])
            
        except Exception as e:
            print(f"文本对齐表格检测错误: {e}")
            traceback.print_exc()
            return None

    def analyze_table_structure(self, page, table):
        """
        分析表格结构（行和列）
        
        参数:
            page: fitz.Page对象
            table: 表格对象，包含bbox
            
        返回:
            更新表格对象，添加行和列信息
        """
        try:
            import fitz
            
            # 获取表格区域
            table_rect = fitz.Rect(table["bbox"])
            
            # 从表格区域提取文本块
            blocks = page.get_text("dict", clip=table_rect)["blocks"]
            
            # 收集所有文本行的Y坐标
            all_lines_y = []
            for block in blocks:
                if block["type"] == 0:  # 文本块
                    for line in block.get("lines", []):
                        y0, y1 = line["bbox"][1], line["bbox"][3]
                        all_lines_y.append(y0)
            
            # 聚类Y坐标以确定行
            y_tolerance = (table_rect.height * 0.01) + 2  # 动态容差
            rows = cluster_positions(all_lines_y, y_tolerance)
            
            # 收集所有spans的X坐标
            all_spans_x = []
            for block in blocks:
                if block["type"] == 0:
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            all_spans_x.append(span["bbox"][0])
            
            # 聚类X坐标以确定列
            x_tolerance = (table_rect.width * 0.01) + 3  # 动态容差
            cols = cluster_positions(all_spans_x, x_tolerance)
            
            # 更新表格对象
            table["rows"] = rows
            table["cols"] = cols
            
            return table
            
        except Exception as e:
            print(f"表格结构分析错误: {e}")
            return table

    def cluster_positions(positions, tolerance=5):
        """
        将接近的位置值聚类
        
        参数:
            positions: 位置值列表
            tolerance: 允许的位置差异
            
        返回:
            聚类后的位置值列表
        """
        if not positions:
            return []
        
        import numpy as np
        
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
    converter.detect_tables_by_grid = types.MethodType(detect_tables_by_grid, converter)
    converter.detect_tables_by_text_alignment = types.MethodType(detect_tables_by_text_alignment, converter)
    converter.analyze_table_structure = types.MethodType(analyze_table_structure, converter)
    
    # 提供全局函数以在模块级别使用
    globals()['cluster_positions'] = cluster_positions
    
    print("已应用增强型表格检测补丁，大幅提高表格识别能力")
    return True

# 如果作为脚本执行，应用补丁到转换器
if __name__ == "__main__":
    try:
        from enhanced_pdf_converter import EnhancedPDFConverter
        converter = EnhancedPDFConverter()
        apply_enhanced_table_detection_patch(converter)
        print("成功应用增强型表格检测补丁到PDF转换器")
    except ImportError:
        print("无法导入EnhancedPDFConverter，请确保文件位于正确的目录")
