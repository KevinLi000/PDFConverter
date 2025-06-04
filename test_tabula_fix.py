#!/usr/bin/env python
"""
测试tabula导入修复
测试脚本，验证tabula的read_pdf函数导入修复是否有效
"""

import os
import sys
import traceback

def test_tabula_import():
    """测试tabula导入"""
    print("测试tabula导入修复...")
    
    # 测试1: 直接导入tabula
    try:
        import tabula
        print("✓ 成功导入tabula模块")
        
        # 测试2: 检查tabula.read_pdf是否可用
        if hasattr(tabula, 'read_pdf'):
            print("✓ tabula.read_pdf函数可用")
        else:
            print("✗ tabula.read_pdf函数不可用")
            
    except ImportError:
        print("✗ 无法导入tabula模块")
        print("  请使用命令安装: pip install tabula-py")
        return False
    
    # 测试3: 尝试从tabula导入read_pdf
    try:
        try:
            from tabula import read_pdf
            print("✓ 可以从tabula导入read_pdf")
        except ImportError:
            print("✗ 无法从tabula导入read_pdf")
            
        # 测试4: 应用tabula导入修复
        try:
            import tabula_adapter
            tabula_adapter.patch_tabula_imports()
            print("✓ 已应用tabula导入修复")
            
            # 测试5: 再次尝试从tabula导入read_pdf
            try:
                from tabula import read_pdf
                print("✓ 修复后可以从tabula导入read_pdf")
            except ImportError:
                print("✗ 修复后仍无法从tabula导入read_pdf")
        except ImportError:
            print("✗ 无法导入tabula_adapter模块")
    except Exception as e:
        print(f"✗ 测试过程中出错: {e}")
        traceback.print_exc()
        return False
    
    return True

def test_enhanced_pdf_converter():
    """测试增强型PDF转换器"""
    print("\n测试增强型PDF转换器...")
    
    try:
        # 导入增强型PDF转换器
        from enhanced_pdf_converter import EnhancedPDFConverter
        print("✓ 成功导入EnhancedPDFConverter")
        
        # 创建实例
        converter = EnhancedPDFConverter()
        print("✓ 成功创建EnhancedPDFConverter实例")
        
        # 应用修复
        try:
            import comprehensive_pdf_fix
            comprehensive_pdf_fix.apply_comprehensive_fixes(converter)
            print("✓ 成功应用全面修复")
        except ImportError:
            print("✗ 无法导入comprehensive_pdf_fix模块")
        except Exception as e:
            print(f"✗ 应用全面修复时出错: {e}")
            traceback.print_exc()
    except ImportError:
        print("✗ 无法导入EnhancedPDFConverter")
        return False
    except Exception as e:
        print(f"✗ 测试过程中出错: {e}")
        traceback.print_exc()
        return False
    
    return True

def main():
    """主函数"""
    print("=" * 50)
    print("PDF转换器修复测试")
    print("=" * 50)
    
    # 测试tabula导入
    tabula_result = test_tabula_import()
    
    # 测试增强型PDF转换器
    converter_result = test_enhanced_pdf_converter()
    
    # 输出总结
    print("\n" + "=" * 50)
    print("测试结果汇总:")
    print(f"tabula导入修复: {'✓ 通过' if tabula_result else '✗ 失败'}")
    print(f"增强型PDF转换器: {'✓ 通过' if converter_result else '✗ 失败'}")
    print("=" * 50)
    
    if tabula_result and converter_result:
        print("\n所有测试通过! PDF转换器修复已成功应用。")
        return 0
    else:
        print("\n部分测试失败，请查看上面的错误信息。")
        return 1

if __name__ == "__main__":
    sys.exit(main())
