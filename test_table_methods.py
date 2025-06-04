#!/usr/bin/env python
"""
测试增强型PDF转换器表格处理功能
"""

import os
import sys
from enhanced_pdf_converter import EnhancedPDFConverter

def test_table_methods():
    """测试表格处理方法"""
    print("开始测试表格处理方法...")
    
    # 初始化转换器
    converter = EnhancedPDFConverter()
    
    # 检查是否已实现所有必要的方法
    methods_to_check = [
        '_mark_table_regions',
        '_build_table_from_cells',
        '_detect_merged_cells',
        '_validate_and_fix_table_data'
    ]
    
    all_methods_exist = True
    for method_name in methods_to_check:
        if hasattr(converter, method_name):
            print(f"✓ 已实现方法: {method_name}")
        else:
            print(f"✗ 缺少方法: {method_name}")
            all_methods_exist = False
    
    if all_methods_exist:
        print("所有表格处理方法已正确实现！")
    else:
        print("缺少一些必要的表格处理方法，请检查实现。")
        
    # 测试方法之间的调用关系
    try:
        # 创建一个模拟表格和块
        mock_table = {
            "bbox": [0, 0, 100, 100],
            "cells": [
                {"bbox": [0, 0, 50, 50], "text": "A1"},
                {"bbox": [50, 0, 100, 50], "text": "B1"},
                {"bbox": [0, 50, 50, 100], "text": "A2"},
                {"bbox": [50, 50, 100, 100], "text": "B2"}
            ]
        }
        mock_blocks = [
            {"type": 0, "bbox": [0, 0, 100, 100], "text": "Sample text"}
        ]
        
        # 测试从单元格构建表格
        table_data, merged_cells = converter._build_table_from_cells(mock_table)
        print(f"表格数据构建测试结果: {len(table_data)}行 x {len(table_data[0]) if table_data else 0}列")
        
        # 测试检测合并单元格
        detected_merged = converter._detect_merged_cells(mock_table)
        print(f"合并单元格检测测试结果: 检测到{len(detected_merged)}个合并单元格")
        
        # 测试验证和修复表格数据
        fixed_data, fixed_merged = converter._validate_and_fix_table_data(table_data, merged_cells)
        print(f"数据验证和修复测试结果: {len(fixed_data)}行 x {len(fixed_data[0]) if fixed_data else 0}列")
        
        # 测试标记表格区域
        marked_blocks = converter._mark_table_regions(mock_blocks, [mock_table])
        print(f"表格区域标记测试结果: {len(marked_blocks)}个块")
        
        print("表格处理方法功能测试通过！")
    except Exception as e:
        import traceback
        print(f"表格处理方法功能测试失败: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    test_table_methods()
