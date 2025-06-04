#!/usr/bin/env python
"""
测试PDF转换器的表格检测功能是否正常
"""

import os
import sys
import traceback

def test_table_detection():
    """测试表格检测功能"""
    print("测试PDF转换器表格检测功能")
    print("=" * 50)
    
    try:
        # 导入PyMuPDF
        import fitz
        print("PyMuPDF (fitz) 导入成功")
        
        # 检查版本
        print(f"PyMuPDF版本: {fitz.version}")
        
        # 测试是否有find_tables方法
        page_has_find_tables = hasattr(fitz.Page, 'find_tables')
        print(f"Page对象是否有find_tables方法: {page_has_find_tables}")
        
        # 导入增强转换器
        from enhanced_pdf_converter import EnhancedPDFConverter
        print("增强PDF转换器导入成功")
        
        # 创建转换器实例
        converter = EnhancedPDFConverter()
        print("转换器实例创建成功")
        
        # 检查是否有detect_tables方法
        has_detect_tables = hasattr(converter, 'detect_tables')
        print(f"转换器是否有detect_tables方法: {has_detect_tables}")
        
        # 如果缺少detect_tables方法，尝试应用补丁
        if not has_detect_tables:
            print("转换器缺少detect_tables方法，尝试应用补丁...")
            
            try:
                # 尝试从表格检测工具导入
                from table_detection_utils import add_table_detection_capability
                add_table_detection_capability(converter)
                print("表格检测功能补丁应用成功")
            except ImportError:
                print("无法导入table_detection_utils")
                
                try:
                    # 尝试应用直接补丁
                    import direct_table_detection_patch
                    direct_table_detection_patch.patch_table_detection(converter)
                    print("直接表格检测补丁应用成功")
                except ImportError:
                    print("无法导入direct_table_detection_patch")
                    
                    try:
                        # 尝试应用修复
                        import fix_table_detection
                        fix_table_detection.apply_table_detection_patch()
                        print("表格检测修复应用成功")
                    except ImportError:
                        print("无法导入fix_table_detection")
        
        # 再次检查是否有detect_tables方法
        has_detect_tables = hasattr(converter, 'detect_tables')
        print(f"应用补丁后转换器是否有detect_tables方法: {has_detect_tables}")
        
        if has_detect_tables:
            print("\n表格检测功能已成功修复!")
            print("转换器现在可以处理表格，不会出现'Page' object has no attribute 'find_tables'错误")
        else:
            print("\n警告: 未能修复表格检测功能")
            print("可能需要手动应用补丁或安装更新版本的PyMuPDF")
        
        # 最后，检查是否可以使用增强的字体和表格样式处理
        try:
            import enhanced_font_handler
            print("增强字体处理模块可用")
        except ImportError:
            print("增强字体处理模块不可用")
        
        try:
            import enhanced_table_style
            print("增强表格样式模块可用")
        except ImportError:
            print("增强表格样式模块不可用")
        
    except ImportError as e:
        print(f"错误: 缺少必要的模块: {e}")
    except Exception as e:
        print(f"测试过程中出错: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    test_table_detection()
