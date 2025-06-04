"""
专门修复PDF转换器中'dict' object has no attribute 'cells'错误
"""

import os
import sys
import types

def apply_dict_cells_fix(converter_instance):
    """
    应用特定的修复来解决'dict' object has no attribute 'cells'错误
    
    参数:
        converter_instance: EnhancedPDFConverter的实例
        
    返回:
        修复后的转换器实例
    """
    # 修复构建表格方法
    def build_table_from_cells_fixed(self, table):
        """
        从单元格数据构建表格结构 - 修复的版本，正确处理字典对象
        
        参数:
            table: 表格对象
            
        返回:
            构建的表格数据和合并单元格信息
        """
        # 检查是否是字典对象
        if isinstance(table, dict):
            # 如果是字典对象，尝试从table_data获取数据
            if "table_data" in table and isinstance(table["table_data"], list):
                return table["table_data"], table.get("merged_cells", [])
            # 检查字典是否包含cells属性
            elif "cells" in table and table["cells"]:
                cells = table["cells"]
            else:
                return [], []
        # 检查是否有cells属性
        elif hasattr(table, 'cells') and table.cells:
            cells = table.cells
        else:
            return [], []
            
        try:
            # 识别行和列的位置
            row_positions = set()
            col_positions = set()
            
            for cell in cells:
                # 收集所有行和列的起始位置
                if isinstance(cell, dict) and "bbox" in cell and len(cell["bbox"]) >= 4:
                    # 处理字典形式的单元格
                    bbox = cell["bbox"]
                    row_positions.add(bbox[1])  # 上边界
                    row_positions.add(bbox[3])  # 下边界
                    col_positions.add(bbox[0])  # 左边界
                    col_positions.add(bbox[2])  # 右边界
                elif len(cell) >= 4:  # 确保单元格有足够的坐标信息
                    row_positions.add(cell[1])  # 上边界
                    row_positions.add(cell[3])  # 下边界
                    col_positions.add(cell[0])  # 左边界
                    col_positions.add(cell[2])  # 右边界
                    
            # 排序位置
            row_positions = sorted(row_positions)
            col_positions = sorted(col_positions)
            
            # 创建空表格
            rows_count = len(row_positions) - 1
            cols_count = len(col_positions) - 1
            
            if rows_count <= 0 or cols_count <= 0:
                return [], []
                
            # 初始化表格和占位标记矩阵
            table_data = [["" for _ in range(cols_count)] for _ in range(rows_count)]
            occupied = [[False for _ in range(cols_count)] for _ in range(rows_count)]
            merged_cells = []  # 存储合并单元格信息: (行开始, 列开始, 行结束, 列结束)
            
            # 为每个单元格创建映射，以便查找其在表格中的位置
            cell_position_map = {}
            
            # 首先识别所有单元格的位置
            for cell in cells:
                # 获取单元格坐标
                if isinstance(cell, dict) and "bbox" in cell and len(cell["bbox"]) >= 4:
                    # 处理字典形式的单元格
                    bbox = cell["bbox"]
                    left, top, right, bottom = bbox[0], bbox[1], bbox[2], bbox[3]
                elif len(cell) >= 4:
                    left, top, right, bottom = cell[0], cell[1], cell[2], cell[3]
                else:
                    continue
                
                # 找出单元格在表格网格中的位置
                row_start = row_positions.index(top) if top in row_positions else -1
                row_end = row_positions.index(bottom) if bottom in row_positions else -1
                col_start = col_positions.index(left) if left in col_positions else -1
                col_end = col_positions.index(right) if right in col_positions else -1
                
                # 跳过无效位置
                if row_start < 0 or row_end <= row_start or col_start < 0 or col_end <= col_start:
                    continue
                
                # 存储单元格位置信息
                cell_key = (left, top, right, bottom)
                cell_position_map[cell_key] = (row_start, col_start, row_end, col_end)
            
            # 然后填充表格内容并识别合并单元格
            for cell in cells:
                # 获取单元格坐标和文本
                if isinstance(cell, dict) and "bbox" in cell:
                    # 处理字典形式的单元格
                    bbox = cell["bbox"]
                    if len(bbox) < 4:
                        continue
                        
                    left, top, right, bottom = bbox[0], bbox[1], bbox[2], bbox[3]
                    cell_text = cell.get("text", "")
                elif len(cell) >= 4:
                    left, top, right, bottom = cell[0], cell[1], cell[2], cell[3]
                    
                    if hasattr(cell, 'text'):
                        cell_text = cell.text
                    elif len(cell) > 4 and isinstance(cell[4], str):
                        cell_text = cell[4]
                    else:
                        cell_text = ""
                else:
                    continue
                
                cell_key = (left, top, right, bottom)
                if cell_key not in cell_position_map:
                    continue
                
                row_start, col_start, row_end, col_end = cell_position_map[cell_key]
                
                # 检查是否为合并单元格
                is_merged = row_end > row_start + 1 or col_end > col_start + 1
                
                if is_merged:
                    # 记录合并单元格信息
                    merged_cells.append((row_start, col_start, row_end - 1, col_end - 1))
                    
                    # 标记所有被合并的单元格为已占用
                    for r in range(row_start, row_end):
                        for c in range(col_start, col_end):
                            occupied[r][c] = True
                    
                    # 只在左上角单元格放置内容
                    table_data[row_start][col_start] = cell_text
                else:
                    # 如果单元格未被占用，放置内容
                    if not occupied[row_start][col_start]:
                        table_data[row_start][col_start] = cell_text
            
            return table_data, merged_cells
            
        except Exception as e:
            print(f"构建表格时出错: {e}")
            return [], []
    
    # 修复检测合并单元格方法
    def detect_merged_cells_fixed(self, table):
        """
        检测表格中的合并单元格 - 修复的版本，正确处理字典对象
        
        参数:
            table: 表格对象
            
        返回:
            合并单元格列表，每个元素为 (行开始, 列开始, 行结束, 列结束)
        """
        merged_cells = []
        
        try:
            # 检查是否是字典对象
            if isinstance(table, dict):
                # 如果是字典对象，直接获取merged_cells字段
                return table.get("merged_cells", [])
                
            # 检查表格结构
            if hasattr(table, 'cells') and table.cells:
                cells = table.cells
                
                # 收集边界
                rows = set()
                cols = set()
                
                for cell in cells:
                    if hasattr(cell, 'bbox') and len(cell.bbox) >= 4:
                        rows.add(cell.bbox[1])  # Top
                        rows.add(cell.bbox[3])  # Bottom
                        cols.add(cell.bbox[0])  # Left
                        cols.add(cell.bbox[2])  # Right
                    elif isinstance(cell, (list, tuple)) and len(cell) >= 4:
                        rows.add(cell[1])  # Top
                        rows.add(cell[3])  # Bottom
                        cols.add(cell[0])  # Left
                        cols.add(cell[2])  # Right
            # 字典格式的表格与cells列表
            elif isinstance(table, dict) and "cells" in table and table["cells"]:
                cells = table["cells"]
                
                # 收集边界
                rows = set()
                cols = set()
                
                for cell in cells:
                    if isinstance(cell, (list, tuple)) and len(cell) >= 4:
                        rows.add(cell[1])  # Top
                        rows.add(cell[3])  # Bottom
                        cols.add(cell[0])  # Left
                        cols.add(cell[2])  # Right
                    elif isinstance(cell, dict) and "bbox" in cell and len(cell["bbox"]) >= 4:
                        bbox = cell["bbox"]
                        rows.add(bbox[1])  # Top
                        rows.add(bbox[3])  # Bottom
                        cols.add(bbox[0])  # Left
                        cols.add(bbox[2])  # Right
            else:
                return []
                
            # 排序边界
            rows = sorted(rows)
            cols = sorted(cols)
            
            # 映射单元格
            for cell in cells:
                cell_bbox = None
                if hasattr(cell, 'bbox') and len(cell.bbox) >= 4:
                    cell_bbox = cell.bbox
                elif isinstance(cell, (list, tuple)) and len(cell) >= 4:
                    cell_bbox = cell
                elif isinstance(cell, dict) and "bbox" in cell and len(cell["bbox"]) >= 4:
                    cell_bbox = cell["bbox"]
                
                if not cell_bbox:
                    continue
                
                # 获取索引
                top_idx = rows.index(cell_bbox[1]) if cell_bbox[1] in rows else -1
                bottom_idx = rows.index(cell_bbox[3]) if cell_bbox[3] in rows else -1
                left_idx = cols.index(cell_bbox[0]) if cell_bbox[0] in cols else -1
                right_idx = cols.index(cell_bbox[2]) if cell_bbox[2] in cols else -1
                
                # 检查合并单元格
                if top_idx >= 0 and bottom_idx > top_idx and left_idx >= 0 and right_idx > left_idx:
                    if bottom_idx - top_idx > 1 or right_idx - left_idx > 1:
                        merged_cells.append((top_idx, left_idx, bottom_idx - 1, right_idx - 1))
            
            # 其他表格类型的备选检测
            if hasattr(table, 'extract'):
                table_data = table.extract()
                if not table_data:
                    return []
                
                rows = len(table_data)
                if rows == 0:
                    return []
                
                cols = len(table_data[0]) if rows > 0 else 0
                if cols == 0:
                    return []
                
                # 跟踪已访问的单元格
                visited = [[False for _ in range(cols)] for _ in range(rows)]
                
                # 检测合并单元格
                for i in range(rows):
                    for j in range(cols):
                        if visited[i][j]:
                            continue
                        
                        current_value = table_data[i][j]
                        visited[i][j] = True
                        
                        # 检查水平合并
                        col_span = 1
                        for c in range(j + 1, cols):
                            if table_data[i][c] == current_value and not visited[i][c]:
                                col_span += 1
                                visited[i][c] = True
                            else:
                                break
                        
                        # 检查垂直合并
                        row_span = 1
                        for r in range(i + 1, rows):
                            valid_range = j + col_span <= cols
                            
                            if valid_range:
                                match = True
                                for c in range(j, j + col_span):
                                    if table_data[r][c] != current_value or visited[r][c]:
                                        match = False
                                        break
                                
                                if match:
                                    row_span += 1
                                    for c in range(j, j + col_span):
                                        visited[r][c] = True
                                else:
                                    break
                            else:
                                break
                        
                        # 记录合并单元格
                        if row_span > 1 or col_span > 1:
                            merged_cells.append((i, j, i + row_span - 1, j + col_span - 1))
        
        except Exception as e:
            print(f"检测合并单元格时出错: {e}")
        
        return merged_cells
    
    # 替换转换器中的方法
    converter_instance._build_table_from_cells = types.MethodType(build_table_from_cells_fixed, converter_instance)
    converter_instance._detect_merged_cells = types.MethodType(detect_merged_cells_fixed, converter_instance)
    
    print("已应用'dict' object has no attribute 'cells'错误修复")
    return converter_instance

# 当脚本被直接运行时的代码
if __name__ == "__main__":
    print("此脚本用于修复PDF转换器中的'dict' object has no attribute 'cells'错误")
    print("请通过导入和调用apply_dict_cells_fix函数来使用")
