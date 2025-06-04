#!/usr/bin/env python
"""
测试PDF转换器修复
测试脚本，验证PDF转换器的各种修复是否有效
"""

import os
import sys
import traceback

def main():
    """主函数"""
    print("=" * 50)
    print("PDF转换器修复测试")
    print("=" * 50)
    
    # 测试1: 检查tabula导入修复
    print("\n1. 测试tabula导入修复")
    try:
        import tabula
        print("  ✓ 成功导入tabula模块")
        
        if hasattr(tabula, 'read_pdf'):
            print("  ✓ tabula.read_pdf函数可用")
        else:
            print("  ✗ tabula.read_pdf函数不可用")
            
        # 应用tabula导入修复
        try:
            import tabula_adapter
            tabula_adapter.patch_tabula_imports()
            print("  ✓ 已应用tabula导入修复")
            
            if hasattr(tabula, 'read_pdf'):
                print("  ✓ 修复后tabula.read_pdf函数可用")
            else:
                print("  ✗ 修复后tabula.read_pdf函数仍不可用")
        except ImportError:
            print("  ✗ 无法导入tabula_adapter模块")
    except ImportError:
        print("  ✗ 无法导入tabula模块")
    
    # 测试2: 测试增强型PDF转换器
    print("\n2. 测试增强型PDF转换器")
    try:
        from enhanced_pdf_converter import EnhancedPDFConverter
        print("  ✓ 成功导入EnhancedPDFConverter")
        
        # 创建实例
        converter = EnhancedPDFConverter()
        print("  ✓ 成功创建EnhancedPDFConverter实例")
        
        # 检查关键方法
        methods_to_check = [
            'pdf_to_word', 
            'pdf_to_excel', 
            'convert_pdf_to_docx', 
            'convert_pdf_to_excel'
        ]
        
        for method in methods_to_check:
            if hasattr(converter, method):
                print(f"  ✓ 方法 {method} 可用")
            else:
                print(f"  ✗ 方法 {method} 不可用")
        
        # 应用方法名称适配
        try:
            import method_name_adapter
            method_name_adapter.apply_method_name_adaptations(converter)
            print("  ✓ 已应用方法名称适配")
            
            # 再次检查关键方法
            all_methods_available = True
            for method in methods_to_check:
                if not hasattr(converter, method):
                    all_methods_available = False
                    print(f"  ✗ 适配后方法 {method} 仍不可用")
            
            if all_methods_available:
                print("  ✓ 适配后所有关键方法都可用")
        except ImportError:
            print("  ✗ 无法导入method_name_adapter模块")
    except ImportError:
        print("  ✗ 无法导入EnhancedPDFConverter")
      # 测试3: 应用全面修复
    print("\n3. 测试全面修复")
    try:
        from enhanced_pdf_converter import EnhancedPDFConverter
        converter = EnhancedPDFConverter()
        
        try:
            import comprehensive_pdf_fix
            comprehensive_pdf_fix.apply_comprehensive_fixes(converter)
            print("  ✓ 成功应用全面修复")
            
            # 检查所有关键方法
            all_methods_available = True
            for method in methods_to_check:
                if not hasattr(converter, method):
                    all_methods_available = False
                    print(f"  ✗ 全面修复后方法 {method} 不可用")
            
            if all_methods_available:
                print("  ✓ 全面修复后所有关键方法都可用")
        except ImportError:
            print("  ✗ 无法导入comprehensive_pdf_fix模块")
        except Exception as e:
            print(f"  ✗ 应用全面修复时出错: {e}")
            traceback.print_exc()
    except ImportError:
        print("  ✗ 无法导入EnhancedPDFConverter")
    
    # 测试4: 测试集成修复程序
    print("\n4. 测试集成修复程序")
    try:
        from enhanced_pdf_converter import EnhancedPDFConverter
        converter = EnhancedPDFConverter()
        
        try:
            import all_pdf_fixes_integrator
            converter = all_pdf_fixes_integrator.integrate_all_fixes(converter)
            print("  ✓ 成功应用所有集成修复")
            
            # 检查表格边框处理
            if hasattr(converter, '_process_table_block'):
                import inspect
                table_method = inspect.getsource(converter._process_table_block)
                if "border_size = 8" in table_method:
                    print("  ✓ 表格边框增强修复已应用")
                else:
                    print("  ✗ 表格边框增强修复未应用")
            else:
                print("  ✗ 找不到_process_table_block方法")
            
            # 检查图像处理
            if hasattr(converter, '_process_image_block_enhanced'):
                import inspect
                image_method = inspect.getsource(converter._process_image_block_enhanced)
                if "extraction_methods" in image_method and "zoom = 4.0" in image_method:
                    print("  ✓ 图像处理增强修复已应用")
                else:
                    print("  ✗ 图像处理增强修复未应用")
            else:
                print("  ✗ 找不到_process_image_block_enhanced方法")
            
            # 检查表格区域标记
            if hasattr(converter, '_mark_table_regions'):
                import inspect
                mark_method = inspect.getsource(converter._mark_table_regions)
                if "comprehensive_pdf_fix" in mark_method:
                    print("  ✓ 表格区域标记整合已应用")
                else:
                    print("  ✗ 表格区域标记整合未应用")
            else:
                print("  ✗ 找不到_mark_table_regions方法")
                
        except ImportError:
            print("  ✗ 无法导入all_pdf_fixes_integrator模块")
        except Exception as e:
            print(f"  ✗ 应用集成修复时出错: {e}")
            traceback.print_exc()
    except ImportError:
        print("  ✗ 无法导入EnhancedPDFConverter")
    
    print("\n" + "=" * 50)
    print("测试完成")
    print("=" * 50)

if __name__ == "__main__":
    main()
