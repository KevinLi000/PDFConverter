"""
PDF转换器表格和图像修复 - 快速应用脚本
"""

import os
import sys
import traceback

def main():
    """应用表格和图像修复并进行测试"""
    print("===== PDF转换器表格和图像修复 =====")
    
    # 检查命令行参数
    if len(sys.argv) > 1:
        action = sys.argv[1].lower()
        
        if action == "apply":
            # 应用修复
            apply_fixes()
        elif action == "test":
            # 测试修复
            test_fixes()
        elif action == "gui":
            # 整合到GUI
            integrate_to_gui()
        elif action == "fix":
            if len(sys.argv) > 3:
                # 修复指定的PDF
                fix_pdf(sys.argv[2], sys.argv[3])
            else:
                print("用法: python apply_fixes.py fix <input.pdf> <output.docx>")
        else:
            show_usage()
    else:
        # 默认执行所有操作
        all_operations()

def show_usage():
    """显示使用说明"""
    print("\n使用方法:")
    print("  python apply_fixes.py [选项]")
    print("\n选项:")
    print("  apply   - 应用表格和图像修复")
    print("  test    - 测试修复效果")
    print("  gui     - 整合修复到GUI")
    print("  fix <input.pdf> <output.docx> - 修复指定的PDF文件")
    print("  不指定选项将执行所有操作")

def all_operations():
    """执行所有操作"""
    apply_fixes()
    integrate_to_gui()
    test_fixes()

def apply_fixes():
    """应用表格和图像修复"""
    print("\n>>> 应用表格和图像修复...")
    try:
        # 导入修复模块
        try:
            from apply_table_image_fixes import apply_all_fixes_to_converter
            
            # 创建转换器实例
            try:
                from enhanced_pdf_converter import EnhancedPDFConverter
                converter = EnhancedPDFConverter()
            except ImportError:
                try:
                    from improved_pdf_converter import ImprovedPDFConverter
                    converter = ImprovedPDFConverter()
                except ImportError:
                    print("错误: 无法创建转换器实例")
                    return
            
            # 应用修复
            converter = apply_all_fixes_to_converter(converter)
            
            if converter:
                print("修复已成功应用")
            else:
                print("应用修复失败")
                
        except ImportError:
            # 尝试导入table_image_fix模块
            try:
                from table_image_fix import apply_table_and_image_fix
                
                # 创建转换器实例
                try:
                    from enhanced_pdf_converter import EnhancedPDFConverter
                    converter = EnhancedPDFConverter()
                except ImportError:
                    try:
                        from improved_pdf_converter import ImprovedPDFConverter
                        converter = ImprovedPDFConverter()
                    except ImportError:
                        print("错误: 无法创建转换器实例")
                        return
                
                # 应用修复
                if apply_table_and_image_fix(converter):
                    print("修复已成功应用")
                else:
                    print("应用修复失败")
                    
            except ImportError:
                print("错误: 无法导入修复模块，请确保修复文件位于正确的目录")
    except Exception as e:
        print(f"应用修复时出错: {e}")
        traceback.print_exc()

def integrate_to_gui():
    """整合修复到GUI"""
    print("\n>>> 整合修复到GUI...")
    try:
        # 导入整合模块
        try:
            from integrate_table_image_fixes_to_gui import integrate_fixes_to_gui
            
            # 执行整合
            if integrate_fixes_to_gui():
                print("修复已成功整合到GUI")
            else:
                print("整合到GUI失败")
                
        except ImportError:
            print("错误: 无法导入整合模块，请确保integrate_table_image_fixes_to_gui.py位于正确的目录")
    except Exception as e:
        print(f"整合到GUI时出错: {e}")
        traceback.print_exc()

def test_fixes():
    """测试修复效果"""
    print("\n>>> 测试修复效果...")
    try:
        # 导入测试模块
        try:
            from test_table_image_fixes import test_table_image_fixes
            
            # 执行测试
            if test_table_image_fixes():
                print("测试通过，修复有效")
            else:
                print("测试失败，修复可能无效")
                
        except ImportError:
            print("错误: 无法导入测试模块，请确保test_table_image_fixes.py位于正确的目录")
    except Exception as e:
        print(f"测试时出错: {e}")
        traceback.print_exc()

def fix_pdf(input_pdf, output_docx):
    """修复指定的PDF文件"""
    print(f"\n>>> 修复PDF文件: {input_pdf}")
    
    if not os.path.exists(input_pdf):
        print(f"错误: 输入文件不存在: {input_pdf}")
        return
    
    try:
        # 尝试应用修复并转换
        try:
            # 先尝试使用apply_all_fixes_to_converter
            try:
                from apply_table_image_fixes import apply_all_fixes_to_converter
                
                # 创建转换器实例
                try:
                    from enhanced_pdf_converter import EnhancedPDFConverter
                    converter = EnhancedPDFConverter()
                except ImportError:
                    try:
                        from improved_pdf_converter import ImprovedPDFConverter
                        converter = ImprovedPDFConverter()
                    except ImportError:
                        raise ImportError("无法创建转换器实例")
                
                # 应用修复
                converter = apply_all_fixes_to_converter(converter)
                
                if not converter:
                    raise Exception("应用修复失败")
                
                # 执行转换
                print(f"正在转换 {input_pdf} 到 {output_docx}...")
                converter.convert_pdf_to_docx(input_pdf, output_docx)
                
            except ImportError:
                # 尝试使用table_image_fix
                from table_image_fix import apply_table_and_image_fix
                
                # 创建转换器实例
                try:
                    from enhanced_pdf_converter import EnhancedPDFConverter
                    converter = EnhancedPDFConverter()
                except ImportError:
                    try:
                        from improved_pdf_converter import ImprovedPDFConverter
                        converter = ImprovedPDFConverter()
                    except ImportError:
                        raise ImportError("无法创建转换器实例")
                
                # 应用修复
                if not apply_table_and_image_fix(converter):
                    raise Exception("应用修复失败")
                
                # 执行转换
                print(f"正在转换 {input_pdf} 到 {output_docx}...")
                converter.convert_pdf_to_docx(input_pdf, output_docx)
            
            # 检查结果
            if os.path.exists(output_docx):
                print(f"转换成功: {output_docx}")
            else:
                print(f"转换失败: 未生成输出文件")
                
        except ImportError as ie:
            print(f"错误: {ie}")
            print("请确保所有必要的文件位于正确的目录")
            
    except Exception as e:
        print(f"修复PDF时出错: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    main()
