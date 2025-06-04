"""
提供标记表格区域和表格单元格处理的功能
"""

import fitz
import traceback

def mark_table_regions(self, blocks, tables):
    """
    标记属于表格区域的块 - 兼容不同表格对象格式
    
    参数:
        blocks: 页面的内容块
        tables: 在页面中检测到的表格列表
    
    返回:
        更新后的块列表，带有表格标记
    """
    if not tables:
        return blocks
        
    # 复制块列表以避免修改原始数据
    marked_blocks = []
    
    # 将表格转换为块并标记
    for table in tables:
        table_rect = None
        
        # 1. 尝试获取表格矩形区域
        try:
            # 方法1: 直接访问rect属性 (PyMuPDF 1.18.0+)
            if hasattr(table, 'rect'):
                table_rect = table.rect
            # 方法2: 直接访问bbox属性
            elif hasattr(table, 'bbox'):
                table_rect = fitz.Rect(table.bbox)
            # 方法3: 字典对象的bbox属性
            elif isinstance(table, dict) and "bbox" in table:
                table_rect = fitz.Rect(table["bbox"])
            # 方法4: 从单元格计算表格范围
            else:
                cells = None
                
                # 获取单元格列表
                if hasattr(table, 'cells') and table.cells:
                    cells = table.cells
                elif isinstance(table, dict) and "cells" in table and table["cells"]:
                    cells = table["cells"]
                elif hasattr(table, 'tables') and table.tables and len(table.tables) > 0:
                    first_table = table.tables[0]
                    if hasattr(first_table, 'cells') and first_table.cells:
                        cells = first_table.cells
                
                if cells and len(cells) > 0:
                    # 从单元格计算表格范围
                    bboxes = []
                    
                    for cell in cells:
                        cell_bbox = None
                        
                        if isinstance(cell, dict) and "bbox" in cell and len(cell["bbox"]) >= 4:
                            cell_bbox = cell["bbox"]
                        elif isinstance(cell, (list, tuple)) and len(cell) >= 4:
                            cell_bbox = cell[:4]
                        elif hasattr(cell, 'bbox') and len(cell.bbox) >= 4:
                            cell_bbox = cell.bbox
                            
                        if cell_bbox:
                            bboxes.append(cell_bbox)
                    
                    if bboxes:
                        min_x = min(bbox[0] for bbox in bboxes)
                        min_y = min(bbox[1] for bbox in bboxes)
                        max_x = max(bbox[2] for bbox in bboxes)
                        max_y = max(bbox[3] for bbox in bboxes)
                        table_rect = fitz.Rect(min_x, min_y, max_x, max_y)
            
            # 如果无法获取表格区域，跳过此表格
            if not table_rect:
                print("警告: 无法获取表格区域，跳过此表格")
                continue
        except Exception as e:
            print(f"警告: 处理表格边界时出错: {e}")
            continue
            
        # 2. 提取并修正表格数据
        try:
            table_data = []
            merged_cells = []
            
            # 方法1: 使用extract方法
            if hasattr(table, 'extract'):
                merged_cells = self._detect_merged_cells(table)
                table_data, _ = self._validate_and_fix_table_data(table.extract(), merged_cells)
            # 方法2: 字典中已包含table_data
            elif isinstance(table, dict) and "table_data" in table:
                table_data = table["table_data"]
                merged_cells = table.get("merged_cells", [])
            # 方法3: 从单元格构建表格
            else:
                table_data, merged_cells = self._build_table_from_cells(table)
        except Exception as e:
            print(f"警告: 提取表格数据时出错: {e}")
            traceback.print_exc()
            table_data = []
            merged_cells = []
        
        # 跳过空表格
        if not table_data:
            continue
            
        # 创建表格块
        table_block = {
            "type": 100,  # 自定义类型表示表格
            "bbox": [table_rect.x0, table_rect.y0, table_rect.x1, table_rect.y1],
            "is_table": True,
            "table_data": table_data,
            "merged_cells": merged_cells,
            "rows": len(table_data),
            "cols": len(table_data[0]) if table_data and table_data[0] else 0
        }
        
        marked_blocks.append(table_block)
    
    # 添加非表格区域的块
    for block in blocks:
        block_rect = fitz.Rect(block["bbox"])
        
        # 检查此块是否与任何表格重叠
        is_in_table = False
        for table_block in [b for b in marked_blocks if b.get("is_table", False)]:
            table_rect = fitz.Rect(table_block["bbox"])
            # 检查重叠
            if block_rect.intersects(table_rect) and block_rect.get_area() / block_rect.get_area() > 0.5:
                is_in_table = True
                break
                
        # 如果不在表格中，添加到最终块列表
        if not is_in_table:
            marked_blocks.append(block)
    
    # 按垂直位置排序
    marked_blocks.sort(key=lambda b: b["bbox"][1])
    return marked_blocks

def build_table_from_cells(self, table):
    """
    从单元格数据构建表格结构
    
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
    # 如果表格有tables属性 (PyMuPDF 1.18.0+)
    elif hasattr(table, 'tables') and table.tables:
        # 尝试获取第一个表格
        if len(table.tables) > 0:
            first_table = table.tables[0]
            if hasattr(first_table, 'cells') and first_table.cells:
                cells = first_table.cells
            else:
                return [], []
        else:
            return [], []
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
            elif isinstance(cell, (list, tuple)) and len(cell) >= 4:  # 确保单元格有足够的坐标信息
                row_positions.add(cell[1])  # 上边界
                row_positions.add(cell[3])  # 下边界
                col_positions.add(cell[0])  # 左边界
                col_positions.add(cell[2])  # 右边界
            elif hasattr(cell, 'bbox') and len(cell.bbox) >= 4:
                row_positions.add(cell.bbox[1])  # 上边界
                row_positions.add(cell.bbox[3])  # 下边界
                col_positions.add(cell.bbox[0])  # 左边界
                col_positions.add(cell.bbox[2])  # 右边界
                
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
            cell_bbox = None
            cell_text = ""
            
            if isinstance(cell, dict) and "bbox" in cell and len(cell["bbox"]) >= 4:
                # 处理字典形式的单元格
                cell_bbox = cell["bbox"]
                cell_text = cell.get("text", "")
            elif isinstance(cell, (list, tuple)) and len(cell) >= 4:
                cell_bbox = cell[:4]
                if len(cell) > 4 and isinstance(cell[4], str):
                    cell_text = cell[4]
            elif hasattr(cell, 'bbox') and len(cell.bbox) >= 4:
                cell_bbox = cell.bbox
                if hasattr(cell, 'text'):
                    cell_text = cell.text
            
            if not cell_bbox:
                continue
                
            left, top, right, bottom = cell_bbox[0], cell_bbox[1], cell_bbox[2], cell_bbox[3]
            
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
            cell_position_map[cell_key] = (row_start, col_start, row_end, col_end, cell_text)
        
        # 然后填充表格内容并识别合并单元格
        for cell_key, position_info in cell_position_map.items():
            row_start, col_start, row_end, col_end, cell_text = position_info
            
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
        traceback.print_exc()
        return [], []

def detect_merged_cells(self, table):
    """
    检测表格中的合并单元格
    
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
            if "merged_cells" in table:
                return table.get("merged_cells", [])
                
            # 如果没有merged_cells字段，尝试从cells分析
            if "cells" not in table or not table["cells"]:
                return []
                
            cells = table["cells"]
        # 如果是表格对象
        elif hasattr(table, 'cells') and table.cells:
            cells = table.cells
        # 如果表格有tables属性 (PyMuPDF 1.18.0+)
        elif hasattr(table, 'tables') and table.tables:
            # 尝试获取第一个表格
            if len(table.tables) > 0:
                first_table = table.tables[0]
                if hasattr(first_table, 'cells') and first_table.cells:
                    cells = first_table.cells
                else:
                    return []
            else:
                return []
        else:
            return []
            
        # 收集边界
        rows = set()
        cols = set()
        
        # 处理不同类型的单元格，提取边界信息
        for cell in cells:
            cell_bbox = None
            
            if isinstance(cell, dict) and "bbox" in cell and len(cell["bbox"]) >= 4:
                cell_bbox = cell["bbox"]
            elif isinstance(cell, (list, tuple)) and len(cell) >= 4:
                cell_bbox = cell[:4]
            elif hasattr(cell, 'bbox') and len(cell.bbox) >= 4:
                cell_bbox = cell.bbox
            
            if not cell_bbox:
                continue
                
            rows.add(cell_bbox[1])  # Top
            rows.add(cell_bbox[3])  # Bottom
            cols.add(cell_bbox[0])  # Left
            cols.add(cell_bbox[2])  # Right
            
        # 排序边界
        rows = sorted(rows)
        cols = sorted(cols)
        
        # 映射单元格
        for cell in cells:
            cell_bbox = None
            
            if isinstance(cell, dict) and "bbox" in cell and len(cell["bbox"]) >= 4:
                cell_bbox = cell["bbox"]
            elif isinstance(cell, (list, tuple)) and len(cell) >= 4:
                cell_bbox = cell[:4]
            elif hasattr(cell, 'bbox') and len(cell.bbox) >= 4:
                cell_bbox = cell.bbox
            
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
        
        # 如果上述方法无法检测到合并单元格，尝试备用方法
        if not merged_cells and hasattr(table, 'extract'):
            try:
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
                            if c < cols and table_data[i][c] == current_value and not visited[i][c]:
                                col_span += 1
                                visited[i][c] = True
                            else:
                                break
                        
                        # 检查垂直合并
                        row_span = 1
                        for r in range(i + 1, rows):
                            if r < rows:
                                valid_range = True
                                for c in range(j, min(j + col_span, cols)):
                                    if c >= cols or r >= rows or table_data[r][c] != current_value or visited[r][c]:
                                        valid_range = False
                                        break
                                
                                if valid_range:
                                    row_span += 1
                                    for c in range(j, min(j + col_span, cols)):
                                        visited[r][c] = True
                                else:
                                    break
                            else:
                                break
                        
                        # 记录合并单元格
                        if row_span > 1 or col_span > 1:
                            merged_cells.append((i, j, i + row_span - 1, j + col_span - 1))
            except Exception as e:
                print(f"备用合并单元格检测失败: {e}")
    
    except Exception as e:
        print(f"检测合并单元格时出错: {e}")
        traceback.print_exc()
    
    return merged_cells
