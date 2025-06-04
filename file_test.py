#!/usr/bin/env python
"""
使用文件日志的简单测试 - 验证增强型PDF转换器表格方法
"""

import sys
from enhanced_pdf_converter import EnhancedPDFConverter

# 重定向输出到文件
with open('test_results.log', 'w') as f:
    # 创建转换器实例
    f.write("Creating converter instance...\n")
    try:
        converter = EnhancedPDFConverter()
        f.write("Converter instance created successfully.\n")
    except Exception as e:
        f.write(f"Error creating converter: {e}\n")
        sys.exit(1)

    # 检查方法是否存在
    f.write("\nTesting if methods exist:\n")
    f.write(f"_mark_table_regions exists: {hasattr(converter, '_mark_table_regions')}\n")
    f.write(f"_build_table_from_cells exists: {hasattr(converter, '_build_table_from_cells')}\n")
    f.write(f"_detect_merged_cells exists: {hasattr(converter, '_detect_merged_cells')}\n")
    f.write(f"_validate_and_fix_table_data exists: {hasattr(converter, '_validate_and_fix_table_data')}\n")
    f.write("All tests complete!\n")
