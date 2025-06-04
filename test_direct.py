#!/usr/bin/env python
"""
直接测试PDF转换器的表格方法
"""

import sys
import traceback

# 直接输出到控制台
def main():
    try:
        print("Testing EnhancedPDFConverter table methods")
        
        from enhanced_pdf_converter import EnhancedPDFConverter
        print("Successfully imported EnhancedPDFConverter")
        
        # 创建实例
        converter = EnhancedPDFConverter()
        print("Successfully created converter instance")
        
        # 检查方法
        method_names = [
            '_mark_table_regions',
            '_build_table_from_cells',
            '_detect_merged_cells',
            '_validate_and_fix_table_data'
        ]
        
        all_methods_present = True
        for method in method_names:
            has_method = hasattr(converter, method)
            print(f"Method {method} exists: {has_method}")
            if not has_method:
                all_methods_present = False
        
        if all_methods_present:
            print("All required methods are present!")
        else:
            print("Some methods are missing!")
            
        return 0
    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
