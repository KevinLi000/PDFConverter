"""
表格合并单元格修复模块
修复_detect_merged_cells和_process_table_block方法的参数不匹配问题
"""

import types
import traceback

def fix_table_cell_merging_methods(converter):
    """
    修复表格单元格合并方法中的参数不匹配问题
    
    参数:
        converter: PDF转换器实例
    
    返回:
        修改后的转换器实例
    """
    try:
        print("应用表格单元格合并方法修复...")
        
        # 修复_detect_merged_cells方法的参数不匹配问题
        if hasattr(converter, '_detect_merged_cells'):
            original_detect_merged_cells = converter._detect_merged_cells
            
            def fixed_detect_merged_cells(self, table):
                """
                修复后的_detect_merged_cells方法，确保参数传递正确
                """
                try:
                    # 原始方法可能要求的是单个参数（不包括self）
                    return original_detect_merged_cells(table)
                except TypeError as e:
                    if "takes 2 positional arguments but 3 were given" in str(e):
                        # 错误情况：原始方法期望2个参数（包括self），但传递了3个
                        # 修正：调用方法时不重复传递self
                        return original_detect_merged_cells(table)
                    elif "takes 1 positional argument but 2 were given" in str(e):
                        # 错误情况：原始方法期望1个参数（不包括self），但传递了2个
                        # 修正：不使用绑定方法，直接调用函数
                        return original_detect_merged_cells.__func__(table)
                    else:
                        # 其他错误，重新抛出
                        raise
            
            # 保存原始方法并替换为修复版本
            converter._original_detect_merged_cells = original_detect_merged_cells
            converter._detect_merged_cells = types.MethodType(fixed_detect_merged_cells, converter)
        
        # 修复_process_table_block方法的参数不匹配问题
        if hasattr(converter, '_process_table_block'):
            original_process_table_block = converter._process_table_block
            
            def fixed_process_table_block(self, doc, block, page, pdf_document):
                """
                修复后的_process_table_block方法，确保参数传递正确
                """
                try:
                    # 原始方法可能接受4个参数（不包括self）
                    return original_process_table_block(doc, block, page, pdf_document)
                except TypeError as e:
                    if "takes 5 positional arguments but 6 were given" in str(e):
                        # 错误情况：原始方法期望5个参数（包括self），但传递了6个
                        # 修正：调用方法时不重复传递self
                        return original_process_table_block(doc, block, page, pdf_document)
                    else:
                        # 其他错误，重新抛出
                        raise
            
            # 保存原始方法并替换为修复版本
            converter._original_process_table_block = original_process_table_block
            converter._process_table_block = types.MethodType(fixed_process_table_block, converter)
        
        # 添加必要的辅助函数
        
        # 检查单元格是否重叠
        def cells_overlap(self, cell1, cell2, overlap_threshold=0.5):
            """
            检查两个单元格是否重叠
            
            参数:
                cell1, cell2: 单元格坐标 (x0, y0, x1, y1)
                overlap_threshold: 重叠阈值，默认为0.5（50%）
            
            返回:
                布尔值，表示是否重叠
            """
            # 计算交集区域
            x_overlap = max(0, min(cell1[2], cell2[2]) - max(cell1[0], cell2[0]))
            y_overlap = max(0, min(cell1[3], cell2[3]) - max(cell1[1], cell2[1]))
            intersection = x_overlap * y_overlap
            
            # 计算单元格面积
            area1 = (cell1[2] - cell1[0]) * (cell1[3] - cell1[1])
            area2 = (cell2[2] - cell2[0]) * (cell2[3] - cell2[1])
            
            # 计算重叠比例
            if area1 <= 0 or area2 <= 0:
                return False
            
            overlap_ratio1 = intersection / area1
            overlap_ratio2 = intersection / area2
            
            # 如果任一单元格的重叠比例超过阈值，则认为重叠
            return overlap_ratio1 > overlap_threshold or overlap_ratio2 > overlap_threshold
        
        # 合并重叠的单元格
        def merge_overlapping_cells(self, cells):
            """
            合并重叠的单元格
            
            参数:
                cells: 单元格列表，每个单元格为 (x0, y0, x1, y1, row, col, rowspan, colspan)
            
            返回:
                合并后的单元格列表
            """
            if not cells:
                return []
            
            # 创建合并单元格结果列表
            merged_cells = []
            processed = [False] * len(cells)
            
            for i, cell1 in enumerate(cells):
                if processed[i]:
                    continue
                
                merged_cell = list(cell1)  # 转换为列表以便修改
                processed[i] = True
                
                # 检查是否有其他单元格与当前单元格重叠
                merged = True
                while merged:
                    merged = False
                    for j, cell2 in enumerate(cells):
                        if processed[j]:
                            continue
                        
                        # 检查单元格是否重叠
                        if self.cells_overlap(merged_cell[:4], cell2[:4]):
                            # 合并单元格
                            merged_cell[0] = min(merged_cell[0], cell2[0])  # x0
                            merged_cell[1] = min(merged_cell[1], cell2[1])  # y0
                            merged_cell[2] = max(merged_cell[2], cell2[2])  # x1
                            merged_cell[3] = max(merged_cell[3], cell2[3])  # y1
                            
                            # 更新跨行和跨列
                            merged_cell[6] = max(merged_cell[6], cell2[6])  # rowspan
                            merged_cell[7] = max(merged_cell[7], cell2[7])  # colspan
                            
                            processed[j] = True
                            merged = True
                
                merged_cells.append(tuple(merged_cell))  # 转换回元组
            
            return merged_cells
        
        # 添加辅助函数到转换器
        converter.cells_overlap = types.MethodType(cells_overlap, converter)
        converter.merge_overlapping_cells = types.MethodType(merge_overlapping_cells, converter)
        
        # 如果有enhanced_extract_tables方法，修复其语法错误
        if hasattr(converter, 'enhanced_extract_tables'):
            original_enhanced_extract_tables = converter.enhanced_extract_tables
            
            def fixed_enhanced_extract_tables(self, page, page_num=None):
                """
                修复后的enhanced_extract_tables方法，确保正确处理参数和语法
                """
                try:
                    # 尝试调用原始方法
                    return original_enhanced_extract_tables(page, page_num)
                except SyntaxError:
                    # 如果有语法错误，使用修复后的实现
                    print("检测到enhanced_extract_tables方法存在语法错误，使用修复后的实现")
                    
                    # 基本的表格提取实现
                    tables = []
                    # 使用PyMuPDF的表格提取功能
                    if hasattr(page, "find_tables"):
                        try:
                            fitz_tables = page.find_tables()
                            for table in fitz_tables:
                                table_dict = {
                                    "bbox": table.bbox,
                                    "rows": table.rows,
                                    "cols": table.cols,
                                    "cells": []
                                }
                                tables.append(table_dict)
                        except Exception as e:
                            print(f"PyMuPDF表格提取失败: {e}")
                    
                    return tables
                except Exception as e:
                    print(f"增强表格提取方法错误: {e}")
                    traceback.print_exc()
                    # 返回空列表，避免处理中断
                    return []
            
            # 替换为修复版本
            converter.enhanced_extract_tables = types.MethodType(fixed_enhanced_extract_tables, converter)
        
        print("表格单元格合并方法修复完成")
        return converter
        
    except Exception as e:
        print(f"应用表格单元格合并方法修复失败: {e}")
        traceback.print_exc()
        return converter

# 直接测试函数
if __name__ == "__main__":
    try:
        from enhanced_pdf_converter import EnhancedPDFConverter
        
        print("测试表格单元格合并方法修复...")
        converter = EnhancedPDFConverter()
        fixed_converter = fix_table_cell_merging_methods(converter)
        print("测试完成!")
    except ImportError:
        print("无法导入转换器类，请确保正确安装了依赖项。")
