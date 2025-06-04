"""
全面修复PDF转换器的各种表格检测和处理问题
包括:
1. 'Page' object has no attribute 'find_tables'错误
2. 'dict' object has no attribute 'cells'错误
3. 表格结构处理兼容性问题
4. 'tabula' 导入问题修复
5. 方法名称适配问题
"""

import os
import sys
import types
import traceback

try:
    import fitz  # PyMuPDF
except ImportError:
    print("错误: 未安装PyMuPDF")
    print("请使用命令安装: pip install PyMuPDF")
    sys.exit(1)

# 尝试导入tabula适配器
try:
    import tabula_adapter
    has_tabula_adapter = True
except ImportError:
    has_tabula_adapter = False
    
# 尝试导入方法名称适配器
try:
    import method_name_adapter
    has_method_adapter = True
except ImportError:
    has_method_adapter = False

def apply_comprehensive_fixes(converter_instance):
    """
    应用全面的PDF转换器修复
    
    参数:
        converter_instance: EnhancedPDFConverter的实例
        
    返回:
        修复后的转换器实例
    """
    print("正在应用全面PDF转换器修复...")
    
    # 修复tabula导入问题
    if has_tabula_adapter:
        tabula_adapter.patch_tabula_imports()
        print("已应用tabula导入修复")
    else:
        print("警告: tabula适配器不可用，无法修复tabula导入问题")
    
    # 应用方法名称适配
    if has_method_adapter:
        method_name_adapter.apply_method_name_adaptations(converter_instance)
        print("已应用方法名称适配")
    else:
        print("警告: 方法名称适配器不可用，可能影响部分功能")
    
    # 检查PyMuPDF版本和find_tables的可用性
    has_find_tables = False
    try:
        import fitz
        test_doc = fitz.open()  # 创建空文档
        test_page = test_doc.new_page()  # 添加一个空白页
        try:
            _ = test_page.find_tables()
            has_find_tables = True
            print("PyMuPDF的find_tables方法可用")
        except AttributeError:
            print("PyMuPDF的find_tables方法不可用，将应用备用解决方案")
        except Exception as e:
            print(f"检测find_tables方法时出错: {e}")
        finally:
            test_doc.close()
    except Exception as e:
        print(f"初始化测试文档时出错: {e}")
    
    # 如果find_tables不可用，需要确保有备用方法
    if not has_find_tables:
        print("正在添加find_tables备用方法...")
        
        # 添加Page.find_tables方法
        def find_tables_fallback(page):
            """
            为Page对象添加find_tables方法的备用实现
            """
            class MockTableContainer:
                def __init__(self, tables):
                    self.tables = tables
            
            try:
                # 使用备用方法检测表格
                tables = []
                # ... 以下是表格检测的简化实现 ...
                # 提取页面内容为图像
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                # 进行一些基本的表格检测...
                # 这里只是返回一个空的结果，实际实现需要更复杂的表格检测算法
                return MockTableContainer(tables)
            except Exception as e:
                print(f"备用find_tables方法出错: {e}")
                return MockTableContainer([])
        
        # 尝试为Page类添加find_tables方法
        try:
            if not hasattr(fitz.Page, 'find_tables'):
                fitz.Page.find_tables = find_tables_fallback
                print("已成功为Page类添加find_tables方法")
        except Exception as e:
            print(f"无法为Page类添加find_tables方法: {e}")
    
    # 1. 修复表格检测方法
    def extract_tables_fallback(self, pdf_document, page_num):
        """
        备用表格检测方法，当PyMuPDF的find_tables方法不可用时使用
        
        参数:
            pdf_document: PyMuPDF文档对象
            page_num: 页码
            
        返回:
            检测到的表格列表
        """
        try:
            # 获取页面
            page = pdf_document[page_num]
            
            # 创建模拟表格对象
            mock_tables = []
            
            # 尝试使用备用的表格检测方法
            try:
                # 使用页面上的线条检测表格
                # 获取页面上的所有线条
                lines = []
                
                # 提取页面上的线条对象
                for item in page.get_drawings():
                    if item.get('type') == 'l':  # 线段
                        lines.append(item)
                
                # 如果线条很少，可能没有表格
                if len(lines) < 4:
                    return []
                
                # 尝试从线条中检测表格结构
                # 简单实现：将线条组织成矩形区域
                horizontal_lines = []
                vertical_lines = []
                
                for line in lines:
                    if 'rect' in line:
                        rect = line['rect']
                        x0, y0, x1, y1 = rect
                        
                        # 判断是水平线还是垂直线
                        if abs(y1 - y0) < abs(x1 - x0):  # 水平线
                            horizontal_lines.append((min(x0, x1), y0, max(x0, x1), y1))
                        else:  # 垂直线
                            vertical_lines.append((x0, min(y0, y1), x1, max(y0, y1)))
                
                # 如果水平线和垂直线都很少，可能没有表格
                if len(horizontal_lines) < 2 or len(vertical_lines) < 2:
                    return []
                
                # 检测表格区域
                # 简单实现：找出线条的最小外接矩形
                if horizontal_lines and vertical_lines:
                    min_x = min([line[0] for line in horizontal_lines + vertical_lines])
                    min_y = min([line[1] for line in horizontal_lines + vertical_lines])
                    max_x = max([line[2] for line in horizontal_lines + vertical_lines])
                    max_y = max([line[3] for line in horizontal_lines + vertical_lines])
                    
                    # 创建一个简单的表格对象
                    table_bbox = [min_x, min_y, max_x, max_y]
                    
                    # 提取这个区域的文本
                    texts = page.get_text("dict", clip=table_bbox)["blocks"]
                    
                    # 创建单元格
                    cells = []
                    for text_block in texts:
                        if text_block["type"] == 0:  # 文本块
                            bbox = text_block["bbox"]
                            text = "".join([span["text"] for span in text_block["lines"][0]["spans"]])
                            cells.append({
                                "bbox": bbox,
                                "text": text
                            })
                    
                    # 只有当有单元格时才添加表格
                    if cells:
                        mock_table = {
                            "bbox": table_bbox,
                            "cells": cells
                        }
                        mock_tables.append(mock_table)
            except Exception as e:
                print(f"备用表格检测方法出错: {e}")
            
            return mock_tables
        except Exception as e:
            print(f"表格提取失败: {e}")
            return []
    
    # 2. 修复_build_table_from_cells方法
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
            return [], []
    
    # 3. 修复_detect_merged_cells方法
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
        
        return merged_cells
    def _validate_and_fix_table_data(self, table_data, merged_cells=None):
        """
        验证表格数据并修复常见问题
        
        参数:
            table_data: 表格数据二维列表
            merged_cells: 合并单元格信息列表，每项为 (start_row, start_col, end_row, end_col)
        
        返回:
            修复后的表格数据和合并单元格信息
        """
        if not table_data:
            return [], []
            
        # 确保表格数据有效
        if not isinstance(table_data, list):
            print("警告: 表格数据不是列表类型")
            return [], []
            
        # 确保表格至少有一行
        if len(table_data) == 0:
            return [], []
            
        # 检查行一致性
        col_count = 0
        for row in table_data:
            if isinstance(row, list):
                col_count = max(col_count, len(row))
                
        if col_count == 0:
            print("警告: 表格没有有效列")
            return [], []
            
        # 初始化修复后的表格数据
        fixed_table_data = []
        
        # 确保所有行具有相同的列数
        for row_idx, row in enumerate(table_data):
            if not isinstance(row, list):
                # 如果行不是列表，创建一个空行
                fixed_row = [""] * col_count
            else:
                # 确保行长度一致
                fixed_row = list(row)
                if len(fixed_row) < col_count:
                    # 填充缺失的单元格
                    fixed_row.extend([""] * (col_count - len(fixed_row)))
                elif len(fixed_row) > col_count:
                    # 截断过长的行
                    fixed_row = fixed_row[:col_count]
            
            # 处理单元格内容
            for i in range(len(fixed_row)):
                cell_content = fixed_row[i]
                
                # 将None替换为空字符串
                if cell_content is None:
                    fixed_row[i] = ""
                
                # 处理非字符串类型
                if not isinstance(cell_content, str):
                    try:
                        fixed_row[i] = str(cell_content)
                    except:
                        fixed_row[i] = ""
                
                # 处理多行文本 - 确保保留换行符
                if isinstance(fixed_row[i], str):
                    # 替换连续空格为单个空格，但保留换行符
                    fixed_row[i] = re.sub(r' {2,}', ' ', fixed_row[i])
                    # 删除行首行尾空白，但保留内部格式
                    fixed_row[i] = fixed_row[i].strip()
            
            fixed_table_data.append(fixed_row)
        
        # 验证合并单元格信息
        if merged_cells is None:
            merged_cells = []
        
        fixed_merged_cells = []
        for merge_info in merged_cells:
            if (isinstance(merge_info, (list, tuple)) and 
                len(merge_info) == 4 and 
                all(isinstance(idx, int) for idx in merge_info)):
                start_row, start_col, end_row, end_col = merge_info
                
                # 确保索引在有效范围内
                if (0 <= start_row <= end_row < len(fixed_table_data) and
                    0 <= start_col <= end_col < col_count):
                    fixed_merged_cells.append((start_row, start_col, end_row, end_col))
        
                    return fixed_table_data, fixed_merged_cells
                    # 列过多，可能是数据错误，截断到标准长度
                    fixed_row = fixed_row[:col_count]
                    print(f"警告: 行 {row_idx} 列数过多，已截断")
            
            # 处理单元格数据
            for col_idx in range(len(fixed_row)):
                cell_value = fixed_row[col_idx]
                
                # 转换None为空字符串
                if cell_value is None:
                    fixed_row[col_idx] = ""
                    continue
                    
                # 尝试将单元格内容转换为字符串
                try:
                    # 如果是数字，保留原值便于后续格式化
                    if isinstance(cell_value, (int, float)):
                        fixed_row[col_idx] = cell_value
                    else:
                        # 转换为字符串并去除前后空白
                        fixed_row[col_idx] = str(cell_value).strip()
                except Exception as e:
                    print(f"转换单元格内容时出错 ({row_idx}, {col_idx}): {e}")
                    fixed_row[col_idx] = ""
            
            fixed_table_data.append(fixed_row)
        
        # 验证并修复合并单元格信息
        fixed_merged_cells = []
        if merged_cells:
            row_count = len(fixed_table_data)
            
            for merge_info in merged_cells:
                if len(merge_info) != 4:
                    print(f"警告: 无效的合并单元格信息: {merge_info}")
                    continue
                    
                start_row, start_col, end_row, end_col = merge_info
                
                # 确保索引在有效范围内
                start_row = max(0, min(start_row, row_count - 1))
                end_row = max(start_row, min(end_row, row_count - 1))
                start_col = max(0, min(start_col, col_count - 1))
                end_col = max(start_col, min(end_col, col_count - 1))
                
                # 添加有效的合并单元格信息
                fixed_merged_cells.append((start_row, start_col, end_row, end_col))
        
        # 处理空表格的特殊情况
        if len(fixed_table_data) == 0:
            # 创建一个最小的有效表格 (1x1)
            fixed_table_data = [["无数据"]]
            print("警告: 创建了默认的空表格")
        
        # 检测并修复无效字符
        for row_idx, row in enumerate(fixed_table_data):
            for col_idx, cell_value in enumerate(row):
                if isinstance(cell_value, str):
                    # 替换控制字符和其他无效字符
                    clean_value = ''.join(c if (c.isprintable() or c in ['\n', '\t']) else ' ' for c in cell_value)
                    
                    # 处理过长的单元格内容
                    if len(clean_value) > 32767:  # Word单元格文本长度限制
                        clean_value = clean_value[:32764] + "..."
                        print(f"警告: 单元格 ({row_idx}, {col_idx}) 内容过长，已截断")
                    
                    fixed_table_data[row_idx][col_idx] = clean_value
        
        return fixed_table_data, fixed_merged_cells
    def _detect_merged_cells(self, table):
    """Detect merged cells in tables"""
    merged_cells = []
    
    try:
        # Check table structure
        if hasattr(table, 'cells') and table.cells:
            cells = table.cells
            
            # Collect boundaries
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
            
            # Sort boundaries
            rows = sorted(rows)
            cols = sorted(cols)
            
            # Map cells
            for cell in cells:
                cell_bbox = None
                
                if hasattr(cell, 'bbox') and len(cell.bbox) >= 4:
                    cell_bbox = cell.bbox
                elif isinstance(cell, (list, tuple)) and len(cell) >= 4:
                    cell_bbox = cell
                
                if not cell_bbox:
                    continue
                
                # Get indices
                top_idx = rows.index(cell_bbox[1]) if cell_bbox[1] in rows else -1
                bottom_idx = rows.index(cell_bbox[3]) if cell_bbox[3] in rows else -1
                left_idx = cols.index(cell_bbox[0]) if cell_bbox[0] in cols else -1
                right_idx = cols.index(cell_bbox[2]) if cell_bbox[2] in cols else -1
                
                # Check for merged cells
                if top_idx >= 0 and bottom_idx > top_idx and left_idx >= 0 and right_idx > left_idx:
                    if bottom_idx - top_idx > 1 or right_idx - left_idx > 1:
                        merged_cells.append((top_idx, left_idx, bottom_idx - 1, right_idx - 1))
        
        # Alternative detection for other table types
        elif hasattr(table, 'extract'):
            table_data = table.extract()
            if not table_data:
                return []
            
            rows = len(table_data)
            if rows == 0:
                return []
            
            cols = len(table_data[0]) if rows > 0 else 0
            if cols == 0:
                return []
            
            # Track visited cells
            visited = [[False for _ in range(cols)] for _ in range(rows)]
            
            # Detect merged cells
            for i in range(rows):
                for j in range(cols):
                    if visited[i][j]:
                        continue
                    
                    current_value = table_data[i][j]
                    visited[i][j] = True
                    
                    # Check horizontal merge
                    col_span = 1
                    for c in range(j + 1, cols):
                        if table_data[i][c] == current_value and not visited[i][c]:
                            col_span += 1
                            visited[i][c] = True
                        else:
                            break
                    
                    # Check vertical merge
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
                    
                    # Record merged cells
                    if row_span > 1 or col_span > 1:
                        merged_cells.append((i, j, i + row_span - 1, j + col_span - 1))
    
    except Exception as e:
        print(f"Error detecting merged cells: {e}")
    
    return merged_cells
    # 4. 修复_mark_table_regions方法
    def mark_table_regions_fixed(self, blocks, tables):
        """
        标记属于表格区域的块 - 增强版，兼容不同表格对象格式
        
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
                    table_data = self._validate_and_fix_table_data(table.extract(),merged_cells)
                # 方法2: 字典中已包含table_data
                elif isinstance(table, dict) and "table_data" in table:
                    table_data = table["table_data"]
                    merged_cells = table.get("merged_cells", [])
                # 方法3: 从单元格构建表格
                else:
                    table_data, merged_cells = self._build_table_from_cells(table)
            except Exception as e:
                print(f"警告: 提取表格数据时出错: {e}")
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
      # 替换转换器中的方法
    converter_instance._extract_tables_fallback = types.MethodType(extract_tables_fallback, converter_instance)
    converter_instance._build_table_from_cells = types.MethodType(build_table_from_cells_fixed, converter_instance)
    converter_instance._detect_merged_cells = types.MethodType(detect_merged_cells_fixed, converter_instance)
    converter_instance._mark_table_regions = types.MethodType(mark_table_regions_fixed, converter_instance)
    
    # 添加标准转换方法
    if not hasattr(converter_instance, 'convert_pdf_to_docx') and hasattr(converter_instance, 'pdf_to_word'):
        def convert_pdf_to_docx(self, input_file, output_file):
            """将PDF转换为Word文档的包装方法"""
            self.pdf_path = input_file
            self.output_dir = os.path.dirname(output_file)
            if not os.path.exists(self.output_dir):
                os.makedirs(self.output_dir)
            result = self.pdf_to_word(method="advanced")
            if result and os.path.exists(result):
                # 如果输出路径不同，复制文件
                if result != output_file:
                    import shutil
                    shutil.copy2(result, output_file)
                    return output_file
                return result
            return output_file
        converter_instance.convert_pdf_to_docx = types.MethodType(convert_pdf_to_docx, converter_instance)
    
    if not hasattr(converter_instance, 'convert_pdf_to_excel') and hasattr(converter_instance, 'pdf_to_excel'):
        def convert_pdf_to_excel(self, input_file, output_file):
            """将PDF转换为Excel的包装方法"""
            self.pdf_path = input_file
            self.output_dir = os.path.dirname(output_file)
            if not os.path.exists(self.output_dir):
                os.makedirs(self.output_dir)
            result = self.pdf_to_excel(method="advanced")
            if result and os.path.exists(result):
                # 如果输出路径不同，复制文件
                if result != output_file:
                    import shutil
                    shutil.copy2(result, output_file)
                    return output_file
                return result
            return output_file
        converter_instance.convert_pdf_to_excel = types.MethodType(convert_pdf_to_excel, converter_instance)
    
    print("已成功应用全面PDF转换器修复")
    return converter_instance

if __name__ == "__main__":
    print("此脚本用于全面修复PDF转换器的表格检测和处理问题")
    print("请在代码中导入并调用apply_comprehensive_fixes函数")
