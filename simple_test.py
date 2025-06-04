#!/usr/bin/env python
"""
简单测试 - 验证增强型PDF转换器表格方法
"""

from enhanced_pdf_converter import EnhancedPDFConverter

# 创建转换器实例
converter = EnhancedPDFConverter()

# 检查方法是否存在
print("Testing if methods exist:")
print(f"_mark_table_regions exists: {hasattr(converter, '_mark_table_regions')}")
print(f"_build_table_from_cells exists: {hasattr(converter, '_build_table_from_cells')}")
print(f"_detect_merged_cells exists: {hasattr(converter, '_detect_merged_cells')}")
print(f"_validate_and_fix_table_data exists: {hasattr(converter, '_validate_and_fix_table_data')}")
print("All tests complete!")
