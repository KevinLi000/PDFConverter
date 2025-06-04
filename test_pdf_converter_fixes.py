#!/usr/bin/env python
"""
测试PDF转换器修复效果
这个脚本测试PDF转换器的表格检测和处理功能
"""

import os
import sys
import traceback

def test_pdf_converter():
    """测试PDF转换器的表格检测和处理功能"""
    print("PDF转换器功能测试")
    print("=" * 50)
    
    try:
        # 1. 导入增强的PDF转换器
        from enhanced_pdf_converter import EnhancedPDFConverter
        
        # 2. 应用修复
        import comprehensive_pdf_fix
        
        # 3. 创建转换器实例
        converter = EnhancedPDFConverter()
        converter = comprehensive_pdf_fix.apply_comprehensive_fixes(converter)
        
        # 4. 测试表格检测功能
        print("\n测试表格检测功能...")
        
        # 检查PyMuPDF版本和find_tables可用性
        import fitz
        print(f"PyMuPDF版本: {fitz.__version__}")
        
        # 测试Page.find_tables的可用性
        doc = fitz.open()  # 创建空文档
        page = doc.new_page()  # 添加空白页面
        
        try:
            tables = page.find_tables()
            print("find_tables方法可用!")
            print(f"  返回值类型: {type(tables)}")
            if hasattr(tables, 'tables'):
                print(f"  tables属性: {type(tables.tables)}")
        except AttributeError:
            print("find_tables方法不可用 - 将使用备用方法")
        except Exception as e:
            print(f"测试find_tables时出错: {e}")
        
        # 5. 测试表格字典对象处理
        print("\n测试表格字典对象处理...")
        
        # 创建模拟的表格字典
        mock_table = {
            "bbox": [100, 100, 300, 200],
            "cells": [
                {"bbox": [100, 100, 200, 150], "text": "单元格1"},
                {"bbox": [200, 100, 300, 150], "text": "单元格2"},
                {"bbox": [100, 150, 200, 200], "text": "单元格3"},
                {"bbox": [200, 150, 300, 200], "text": "单元格4"}
            ]
        }
        
        # 测试_build_table_from_cells方法
        try:
            table_data, merged_cells = converter._build_table_from_cells(mock_table)
            print("_build_table_from_cells方法成功处理字典表格!")
            print(f"  生成的表格数据: {table_data}")
            print(f"  检测到的合并单元格: {merged_cells}")
        except Exception as e:
            print(f"_build_table_from_cells方法处理字典表格失败: {e}")
            traceback.print_exc()
        
        # 测试_detect_merged_cells方法
        try:
            merged_cells = converter._detect_merged_cells(mock_table)
            print("_detect_merged_cells方法成功处理字典表格!")
            print(f"  检测到的合并单元格: {merged_cells}")
        except Exception as e:
            print(f"_detect_merged_cells方法处理字典表格失败: {e}")
            traceback.print_exc()
        
        # 6. 测试_mark_table_regions方法
        print("\n测试_mark_table_regions方法...")
        
        # 创建模拟的内容块
        mock_blocks = [
            {"type": 0, "bbox": [50, 50, 350, 250]},
            {"type": 0, "bbox": [50, 300, 350, 400]}
        ]
        
        # 测试_mark_table_regions方法
        try:
            marked_blocks = converter._mark_table_regions(mock_blocks, [mock_table])
            print("_mark_table_regions方法成功!")
            print(f"  标记后的块数量: {len(marked_blocks)}")
            table_blocks = [b for b in marked_blocks if b.get("is_table", False)]
            print(f"  表格块数量: {len(table_blocks)}")
        except Exception as e:
            print(f"_mark_table_regions方法失败: {e}")
            traceback.print_exc()
        
        print("\n测试完成!")
        return 0
    except Exception as e:
        print(f"测试时出现错误: {e}")
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(test_pdf_converter())
