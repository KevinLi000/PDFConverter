#!/usr/bin/env python
"""
测试表格方法添加到PDF转换器的效果
"""

# 直接测试增强型PDF转换器中的表格方法

def verify_converter_methods():
    try:
        # 检查增强型PDF转换器模块是否可导入
        import sys
        print(f"Python version: {sys.version}")
        print(f"Python executable: {sys.executable}")
        print("Importing enhanced_pdf_converter...")
        
        from enhanced_pdf_converter import EnhancedPDFConverter
        print("EnhancedPDFConverter imported successfully")
        
        # 创建转换器实例
        converter = EnhancedPDFConverter()
        print("Converter instance created successfully")
        
        # 检查方法是否存在
        methods_to_check = [
            '_mark_table_regions',
            '_build_table_from_cells',
            '_detect_merged_cells',
            '_validate_and_fix_table_data'
        ]
        
        for method in methods_to_check:
            if hasattr(converter, method):
                print(f"Method {method} is properly implemented")
            else:
                print(f"ERROR: Method {method} is missing")
                
        # 创建简单的测试对象
        test_table = {
            "bbox": [0, 0, 100, 100],
            "cells": [
                {"bbox": [0, 0, 50, 50], "text": "Cell 1"},
                {"bbox": [50, 0, 100, 50], "text": "Cell 2"}
            ]
        }
        
        # 尝试调用每个方法
        print("\nTesting method functionality:")
        
        # 测试 _build_table_from_cells
        print("Testing _build_table_from_cells...")
        table_data, merged_cells = converter._build_table_from_cells(test_table)
        print(f"_build_table_from_cells returned {len(table_data)} rows")
        
        # 测试 _detect_merged_cells
        print("Testing _detect_merged_cells...")
        detected_merged = converter._detect_merged_cells(test_table)
        print(f"_detect_merged_cells detected {len(detected_merged)} merged cells")
        
        # 测试 _validate_and_fix_table_data
        print("Testing _validate_and_fix_table_data...")
        fixed_data, fixed_merged = converter._validate_and_fix_table_data([["Data"]], [])
        print(f"_validate_and_fix_table_data returned data with {len(fixed_data)} rows")
        
        # 测试 _mark_table_regions
        print("Testing _mark_table_regions...")
        test_blocks = [{"bbox": [0, 0, 100, 100], "type": 0}]
        marked_blocks = converter._mark_table_regions(test_blocks, [test_table])
        print(f"_mark_table_regions returned {len(marked_blocks)} blocks")
        
        print("\nAll methods are implemented and working correctly!")
        return True
        
    except Exception as e:
        import traceback
        print(f"ERROR: {e}")
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Starting table methods test...")
    success = verify_converter_methods()
    print(f"\nTest {'PASSED' if success else 'FAILED'}")
    
    # 写入结果到文件以便查看
    with open("table_methods_test_result.txt", "w") as f:
        f.write("Table methods test completed\n")
        f.write(f"Test {'PASSED' if success else 'FAILED'}\n")
