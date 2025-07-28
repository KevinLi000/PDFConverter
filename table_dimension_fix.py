"""
表格行宽高识别错误修复模块
修复表格维度检测中的精度问题和错误
"""

import os
import sys
import types
import traceback
import fitz
import numpy as np
from docx.shared import Pt, Inches, RGBColor, Cm, Twips
from docx.oxml.shared import OxmlElement, qn
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

def apply_table_dimension_fix(converter):
    """
    应用表格行宽高识别修复
    """
    print("正在应用表格行宽高识别修复...")
    
    def enhanced_dimension_detection(self, table_block, page):
        """
        增强的表格维度检测算法
        解决行高、列宽识别错误的问题
        """
        try:
            dimension_info = {
                "precise_row_heights": [],
                "precise_col_widths": [],
                "cell_coordinates": [],
                "table_grid": None,
                "detection_method": "unknown",
                "confidence_score": 0.0
            }
            
            # 获取表格基本信息
            table_bbox = table_block.get("bbox", [0, 0, 100, 100])
            table_data = table_block.get("table_data", [])
            
            if not table_data:
                return dimension_info
            
            # 方法1: 基于文本块精确分析
            text_based_dims = self._detect_dimensions_from_text_blocks(page, table_bbox, table_data)
            if text_based_dims["confidence_score"] > 0.8:
                dimension_info.update(text_based_dims)
                dimension_info["detection_method"] = "text_blocks"
                return dimension_info
            
            # 方法2: 基于图像分析的边界检测
            image_based_dims = self._detect_dimensions_from_image(page, table_bbox, table_data)
            if image_based_dims["confidence_score"] > 0.7:
                dimension_info.update(image_based_dims)
                dimension_info["detection_method"] = "image_analysis"
                return dimension_info
            
            # 方法3: 基于绘图对象的边界线检测
            drawing_based_dims = self._detect_dimensions_from_drawings(page, table_bbox, table_data)
            if drawing_based_dims["confidence_score"] > 0.6:
                dimension_info.update(drawing_based_dims)
                dimension_info["detection_method"] = "drawing_lines"
                return dimension_info
            
            # 方法4: 智能估算（最后备选）
            estimated_dims = self._estimate_dimensions_intelligently(table_bbox, table_data)
            dimension_info.update(estimated_dims)
            dimension_info["detection_method"] = "intelligent_estimation"
            
            return dimension_info
            
        except Exception as e:
            print(f"表格维度检测出错: {e}")
            traceback.print_exc()
            return self._get_fallback_dimensions(table_bbox, table_data)
    
    def detect_dimensions_from_text_blocks(self, page, table_bbox, table_data):
        """
        基于文本块精确分析表格维度
        """
        try:
            # 获取表格区域内的所有文本块
            table_rect = fitz.Rect(table_bbox)
            text_dict = page.get_text("dict", clip=table_rect)
            
            # 收集所有文本块的位置信息
            text_blocks = []
            for block in text_dict.get("blocks", []):
                if block.get("type") == 0:  # 文本块
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            if span.get("text", "").strip():
                                text_blocks.append({
                                    "bbox": span["bbox"],
                                    "text": span["text"],
                                    "font_size": span.get("size", 10)
                                })
            
            if not text_blocks:
                return {"confidence_score": 0.0}
            
            # 按Y坐标分组确定行
            row_groups = self._group_by_y_coordinate(text_blocks, tolerance=3.0)
            
            # 按X坐标分组确定列
            col_groups = self._group_by_x_coordinate(text_blocks, tolerance=5.0)
            
            # 计算精确的行高
            row_heights = []
            if len(row_groups) > 1:
                sorted_row_positions = sorted(row_groups.keys())
                for i in range(len(sorted_row_positions) - 1):
                    current_y = sorted_row_positions[i]
                    next_y = sorted_row_positions[i + 1]
                    
                    # 计算当前行的底部位置
                    current_row_blocks = row_groups[current_y]
                    current_bottom = max(block["bbox"][3] for block in current_row_blocks)
                    
                    # 计算下一行的顶部位置
                    next_row_blocks = row_groups[next_y]
                    next_top = min(block["bbox"][1] for block in next_row_blocks)
                    
                    # 行高 = 下一行顶部 - 当前行顶部
                    row_height = next_y - current_y
                    row_heights.append(max(row_height, 12))  # 最小行高12点
                
                # 最后一行的高度
                last_row_blocks = row_groups[sorted_row_positions[-1]]
                last_row_height = max(block["font_size"] * 1.5 for block in last_row_blocks)
                row_heights.append(last_row_height)
            
            # 计算精确的列宽
            col_widths = []
            if len(col_groups) > 1:
                sorted_col_positions = sorted(col_groups.keys())
                table_width = table_bbox[2] - table_bbox[0]
                
                for i in range(len(sorted_col_positions) - 1):
                    current_x = sorted_col_positions[i]
                    next_x = sorted_col_positions[i + 1]
                    col_width = next_x - current_x
                    col_widths.append(col_width)
                
                # 最后一列的宽度
                last_col_width = table_bbox[2] - sorted_col_positions[-1]
                col_widths.append(last_col_width)
                
                # 归一化列宽
                total_width = sum(col_widths)
                if total_width > 0:
                    col_widths = [w / total_width for w in col_widths]
            
            # 生成单元格坐标网格
            cell_coordinates = self._generate_cell_coordinates(
                table_bbox, row_groups, col_groups, len(table_data), len(table_data[0]) if table_data else 0
            )
            
            # 计算置信度
            confidence = self._calculate_confidence(len(row_groups), len(col_groups), len(table_data), 
                                                   len(table_data[0]) if table_data else 0)
            
            return {
                "precise_row_heights": row_heights,
                "precise_col_widths": col_widths,
                "cell_coordinates": cell_coordinates,
                "confidence_score": confidence,
                "row_count_detected": len(row_groups),
                "col_count_detected": len(col_groups)
            }
            
        except Exception as e:
            print(f"文本块维度检测出错: {e}")
            return {"confidence_score": 0.0}
    
    def detect_dimensions_from_image(self, page, table_bbox, table_data):
        """
        基于图像分析检测表格维度
        """
        try:
            import cv2
            
            # 高分辨率渲染表格区域
            table_rect = fitz.Rect(table_bbox)
            zoom = 4.0  # 更高的缩放比例提高检测精度
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, clip=table_rect)
            
            # 转换为numpy数组
            img_data = pix.samples
            width, height = pix.width, pix.height
            img_array = np.frombuffer(img_data, dtype=np.uint8)
            img_array = img_array.reshape(height, width, -1)
            
            # 转换为灰度图
            if img_array.shape[2] >= 3:
                gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
            else:
                gray = img_array[:, :, 0]
            
            # 自适应阈值处理
            binary = cv2.adaptiveThreshold(
                gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2
            )
            
            # 检测水平线（行分隔线）
            horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (width // 10, 1))
            horizontal_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel)
            
            # 检测垂直线（列分隔线）
            vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, height // 10))
            vertical_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel)
            
            # 提取水平线位置（行边界）
            h_lines = cv2.HoughLinesP(horizontal_lines, 1, np.pi/180, threshold=width//4, 
                                     minLineLength=width//3, maxLineGap=10)
            
            # 提取垂直线位置（列边界）
            v_lines = cv2.HoughLinesP(vertical_lines, 1, np.pi/180, threshold=height//4, 
                                     minLineLength=height//3, maxLineGap=10)
            
            # 处理水平线位置
            row_positions = []
            if h_lines is not None:
                for line in h_lines:
                    x1, y1, x2, y2 = line[0]
                    # 转换回PDF坐标
                    pdf_y = table_bbox[1] + (y1 / zoom)
                    row_positions.append(pdf_y)
            
            # 处理垂直线位置
            col_positions = []
            if v_lines is not None:
                for line in v_lines:
                    x1, y1, x2, y2 = line[0]
                    # 转换回PDF坐标
                    pdf_x = table_bbox[0] + (x1 / zoom)
                    col_positions.append(pdf_x)
            
            # 排序并去重
            row_positions = sorted(list(set([round(pos, 1) for pos in row_positions])))
            col_positions = sorted(list(set([round(pos, 1) for pos in col_positions])))
            
            # 计算行高
            row_heights = []
            if len(row_positions) > 1:
                for i in range(len(row_positions) - 1):
                    height = row_positions[i + 1] - row_positions[i]
                    row_heights.append(max(height, 8))  # 最小行高
            
            # 计算列宽
            col_widths = []
            if len(col_positions) > 1:
                table_width = table_bbox[2] - table_bbox[0]
                for i in range(len(col_positions) - 1):
                    width = col_positions[i + 1] - col_positions[i]
                    col_widths.append(width / table_width)
            
            # 计算置信度
            expected_rows = len(table_data)
            expected_cols = len(table_data[0]) if table_data else 0
            confidence = self._calculate_confidence(len(row_positions) - 1, len(col_positions) - 1, 
                                                   expected_rows, expected_cols)
            
            return {
                "precise_row_heights": row_heights,
                "precise_col_widths": col_widths,
                "confidence_score": confidence,
                "detected_h_lines": len(row_positions) - 1,
                "detected_v_lines": len(col_positions) - 1
            }
            
        except Exception as e:
            print(f"图像维度检测出错: {e}")
            return {"confidence_score": 0.0}
    
    def detect_dimensions_from_drawings(self, page, table_bbox, table_data):
        """
        基于绘图对象检测表格维度
        """
        try:
            # 获取页面上的绘图对象
            drawings = page.get_drawings()
            
            table_rect = fitz.Rect(table_bbox)
            
            # 收集表格区域内的线条
            h_lines = []  # 水平线
            v_lines = []  # 垂直线
            
            for drawing in drawings:
                for item in drawing.get("items", []):
                    if item[0] == "l":  # 线条
                        x1, y1, x2, y2 = item[1]
                        
                        # 检查线条是否在表格区域内
                        if (table_rect.contains(fitz.Point(x1, y1)) or 
                            table_rect.contains(fitz.Point(x2, y2))):
                            
                            # 判断是水平线还是垂直线
                            if abs(y2 - y1) < 2:  # 水平线
                                h_lines.append((x1, y1, x2, y2))
                            elif abs(x2 - x1) < 2:  # 垂直线
                                v_lines.append((x1, y1, x2, y2))
            
            # 提取行位置
            row_positions = []
            for line in h_lines:
                y_pos = (line[1] + line[3]) / 2
                row_positions.append(y_pos)
            
            # 提取列位置
            col_positions = []
            for line in v_lines:
                x_pos = (line[0] + line[2]) / 2
                col_positions.append(x_pos)
            
            # 排序并去重
            row_positions = sorted(list(set([round(pos, 1) for pos in row_positions])))
            col_positions = sorted(list(set([round(pos, 1) for pos in col_positions])))
            
            # 计算尺寸
            row_heights = []
            if len(row_positions) > 1:
                for i in range(len(row_positions) - 1):
                    height = row_positions[i + 1] - row_positions[i]
                    row_heights.append(max(height, 10))
            
            col_widths = []
            if len(col_positions) > 1:
                table_width = table_bbox[2] - table_bbox[0]
                for i in range(len(col_positions) - 1):
                    width = col_positions[i + 1] - col_positions[i]
                    col_widths.append(width / table_width)
            
            # 计算置信度
            expected_rows = len(table_data)
            expected_cols = len(table_data[0]) if table_data else 0
            confidence = self._calculate_confidence(len(row_positions) - 1, len(col_positions) - 1, 
                                                   expected_rows, expected_cols)
            
            return {
                "precise_row_heights": row_heights,
                "precise_col_widths": col_widths,
                "confidence_score": confidence,
                "h_lines_found": len(h_lines),
                "v_lines_found": len(v_lines)
            }
            
        except Exception as e:
            print(f"绘图对象维度检测出错: {e}")
            return {"confidence_score": 0.0}
    
    def estimate_dimensions_intelligently(self, table_bbox, table_data):
        """
        智能估算表格维度（最后备选方案）
        """
        try:
            rows_count = len(table_data)
            cols_count = len(table_data[0]) if table_data else 0
            
            table_width = table_bbox[2] - table_bbox[0]
            table_height = table_bbox[3] - table_bbox[1]
            
            # 基于内容长度智能分配列宽
            col_widths = []
            if cols_count > 0:
                # 计算每列的文本长度
                col_text_lengths = [0] * cols_count
                for row in table_data:
                    for i, cell in enumerate(row):
                        if i < cols_count and cell is not None:
                            # 考虑中文字符宽度
                            text = str(cell)
                            length = sum(2 if ord(char) > 127 else 1 for char in text)
                            col_text_lengths[i] += length
                
                # 归一化列宽，但设置最小和最大宽度限制
                total_length = sum(col_text_lengths) if sum(col_text_lengths) > 0 else cols_count
                col_widths = []
                for length in col_text_lengths:
                    ratio = max(0.08, min(0.5, length / total_length))  # 限制在8%-50%之间
                    col_widths.append(ratio)
                
                # 确保总和为1
                sum_widths = sum(col_widths)
                if sum_widths > 0:
                    col_widths = [w / sum_widths for w in col_widths]
            
            # 基于内容行数智能分配行高
            row_heights = []
            if rows_count > 0:
                for i, row in enumerate(table_data):
                    # 检查每行内容的复杂度
                    max_lines = 1
                    for cell in row:
                        if cell is not None:
                            text = str(cell)
                            # 估算文本可能的行数
                            estimated_lines = max(1, len(text) // 50)  # 假设每行50字符
                            max_lines = max(max_lines, estimated_lines)
                    
                    # 基础行高 + 内容复杂度调整
                    base_height = 18  # 基础行高
                    content_height = base_height * max_lines
                    
                    # 表头通常高一些
                    if i == 0 and rows_count > 1:
                        content_height *= 1.2
                    
                    row_heights.append(content_height)
            
            return {
                "precise_row_heights": row_heights,
                "precise_col_widths": col_widths,
                "confidence_score": 0.5,  # 中等置信度
                "estimation_method": "content_based"
            }
            
        except Exception as e:
            print(f"智能估算出错: {e}")
            return self._get_fallback_dimensions(table_bbox, table_data)
    
    # 辅助方法
    def group_by_y_coordinate(self, text_blocks, tolerance=3.0):
        """按Y坐标分组文本块"""
        groups = {}
        for block in text_blocks:
            y_coord = block["bbox"][1]
            
            # 查找最接近的组
            found_group = False
            for group_y in groups.keys():
                if abs(group_y - y_coord) <= tolerance:
                    groups[group_y].append(block)
                    found_group = True
                    break
            
            if not found_group:
                groups[y_coord] = [block]
        
        return groups
    
    def group_by_x_coordinate(self, text_blocks, tolerance=5.0):
        """按X坐标分组文本块"""
        groups = {}
        for block in text_blocks:
            x_coord = block["bbox"][0]
            
            # 查找最接近的组
            found_group = False
            for group_x in groups.keys():
                if abs(group_x - x_coord) <= tolerance:
                    groups[group_x].append(block)
                    found_group = True
                    break
            
            if not found_group:
                groups[x_coord] = [block]
        
        return groups
    
    def generate_cell_coordinates(self, table_bbox, row_groups, col_groups, expected_rows, expected_cols):
        """生成单元格坐标网格"""
        try:
            row_positions = sorted(row_groups.keys())
            col_positions = sorted(col_groups.keys())
            
            coordinates = []
            for i in range(expected_rows):
                row_coords = []
                for j in range(expected_cols):
                    if i < len(row_positions) and j < len(col_positions):
                        # 计算单元格边界
                        top = row_positions[i] if i < len(row_positions) else table_bbox[1]
                        bottom = row_positions[i + 1] if i + 1 < len(row_positions) else table_bbox[3]
                        left = col_positions[j] if j < len(col_positions) else table_bbox[0]
                        right = col_positions[j + 1] if j + 1 < len(col_positions) else table_bbox[2]
                        
                        row_coords.append((left, top, right, bottom))
                    else:
                        # 使用估算位置
                        cell_width = (table_bbox[2] - table_bbox[0]) / expected_cols
                        cell_height = (table_bbox[3] - table_bbox[1]) / expected_rows
                        
                        left = table_bbox[0] + j * cell_width
                        right = left + cell_width
                        top = table_bbox[1] + i * cell_height
                        bottom = top + cell_height
                        
                        row_coords.append((left, top, right, bottom))
                
                coordinates.append(row_coords)
            
            return coordinates
            
        except Exception as e:
            print(f"生成单元格坐标出错: {e}")
            return []
    
    def calculate_confidence(self, detected_rows, detected_cols, expected_rows, expected_cols):
        """计算检测置信度"""
        if expected_rows == 0 or expected_cols == 0:
            return 0.0
        
        row_accuracy = 1.0 - abs(detected_rows - expected_rows) / max(detected_rows, expected_rows)
        col_accuracy = 1.0 - abs(detected_cols - expected_cols) / max(detected_cols, expected_cols)
        
        return (row_accuracy + col_accuracy) / 2
    
    def get_fallback_dimensions(self, table_bbox, table_data):
        """获取备用维度信息"""
        rows_count = len(table_data)
        cols_count = len(table_data[0]) if table_data else 0
        
        return {
            "precise_row_heights": [20] * rows_count,  # 默认行高20点
            "precise_col_widths": [1.0 / cols_count] * cols_count if cols_count > 0 else [],
            "confidence_score": 0.3,
            "detection_method": "fallback"
        }
    
    # 应用精确的表格尺寸
    def apply_precise_table_dimensions(self, table, dimension_info):
        """
        应用精确的表格维度到Word表格
        """
        try:
            row_heights = dimension_info.get("precise_row_heights", [])
            col_widths = dimension_info.get("precise_col_widths", [])
            
            # 应用行高
            if row_heights and len(row_heights) == len(table.rows):
                for i, row in enumerate(table.rows):
                    if i < len(row_heights):
                        height_pt = max(12, row_heights[i])  # 最小12点
                        row.height = Pt(height_pt)
                        
                        # 设置行高规则为精确高度
                        tr_pr = row._element.get_or_add_trPr()
                        height_xml = f'''<w:trHeight {nsdecls("w")} w:val="{int(height_pt * 20)}" w:hRule="exact"/>'''
                        
                        # 移除现有高度设置
                        existing_heights = tr_pr.xpath('./w:trHeight')
                        for height_elem in existing_heights:
                            tr_pr.remove(height_elem)
                        
                        # 添加新高度设置
                        tr_pr.append(parse_xml(height_xml))
            
            # 应用列宽
            if col_widths and len(col_widths) == len(table.columns):
                # 获取总宽度
                try:
                    section = table._parent.part.document.sections[0]
                    total_width = section.page_width - section.left_margin - section.right_margin
                    total_width_twips = total_width.twips
                except:
                    total_width_twips = 9000  # 默认宽度
                
                for i, width_ratio in enumerate(col_widths):
                    if i < len(table.columns):
                        column_width_twips = int(total_width_twips * width_ratio)
                        column_width_twips = max(200, column_width_twips)  # 最小200 twips
                        
                        # 设置列宽
                        for cell in table.columns[i].cells:
                            tc_pr = cell._element.get_or_add_tcPr()
                            
                            width_xml = f'<w:tcW {nsdecls("w")} w:w="{column_width_twips}" w:type="dxa"/>'
                            
                            # 移除现有宽度设置
                            existing_width = tc_pr.xpath('./w:tcW')
                            for width_elem in existing_width:
                                tc_pr.remove(width_elem)
                            
                            # 添加新宽度
                            tc_pr.append(parse_xml(width_xml))
            
            print(f"成功应用表格维度 - 行高: {len(row_heights)}, 列宽: {len(col_widths)}")
            
        except Exception as e:
            print(f"应用表格维度出错: {e}")
            traceback.print_exc()
    
    # 将新方法绑定到转换器
    converter._enhanced_dimension_detection = types.MethodType(enhanced_dimension_detection, converter)
    converter._detect_dimensions_from_text_blocks = types.MethodType(detect_dimensions_from_text_blocks, converter)
    converter._detect_dimensions_from_image = types.MethodType(detect_dimensions_from_image, converter)
    converter._detect_dimensions_from_drawings = types.MethodType(detect_dimensions_from_drawings, converter)
    converter._estimate_dimensions_intelligently = types.MethodType(estimate_dimensions_intelligently, converter)
    converter._group_by_y_coordinate = types.MethodType(group_by_y_coordinate, converter)
    converter._group_by_x_coordinate = types.MethodType(group_by_x_coordinate, converter)
    converter._generate_cell_coordinates = types.MethodType(generate_cell_coordinates, converter)
    converter._calculate_confidence = types.MethodType(calculate_confidence, converter)
    converter._get_fallback_dimensions = types.MethodType(get_fallback_dimensions, converter)
    converter.apply_precise_table_dimensions = types.MethodType(apply_precise_table_dimensions, converter)
    
    print("表格行宽高识别修复已成功应用")
    return True

if __name__ == "__main__":
    print("表格维度修复模块已加载")
