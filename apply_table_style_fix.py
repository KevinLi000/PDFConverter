"""
应用表格样式继承修复 - 整合到PDF转换流程
"""

import os
import sys
import types
import traceback

def apply_table_style_fix():
    """应用表格样式继承修复并集成到PDF转换流程"""
    print("开始应用表格样式继承修复...")
    
    # 导入表格样式修复模块
    try:
        from table_style_inheritance_fix import apply_table_style_fixes
    except ImportError:
        print("错误: 无法导入表格样式继承修复模块，请确保文件位于正确的目录")
        return False
    
    # 获取转换器实例
    converter = None
    
    # 尝试导入并创建转换器实例
    try:
        from enhanced_pdf_converter import EnhancedPDFConverter
        converter = EnhancedPDFConverter()
        print("已创建EnhancedPDFConverter实例")
    except ImportError:
        try:
            from improved_pdf_converter import ImprovedPDFConverter
            converter = ImprovedPDFConverter()
            print("已创建ImprovedPDFConverter实例")
        except ImportError:
            print("错误: 无法创建PDF转换器实例，请确保相关文件位于正确的目录")
            return False
    
    # 应用表格样式修复
    if not apply_table_style_fixes(converter):
        print("错误: 应用表格样式修复失败")
        return False
    
    # 整合到PDF转换流程
    integrate_to_conversion_workflow(converter)
    
    print("表格样式继承修复已成功应用并整合到转换流程")
    
    # 测试修复是否正常工作
    test_files = find_test_pdf_files()
    if test_files:
        print("\n您可以测试修复效果，使用以下命令：")
        for test_file in test_files[:2]:  # 最多显示2个测试文件
            output_file = os.path.splitext(test_file)[0] + "_fixed.docx"
            print(f"python3 apply_table_style_fix.py test \"{test_file}\" \"{output_file}\"")
    
    return True

def integrate_to_conversion_workflow(converter):
    """整合表格样式修复到PDF转换流程"""
    # 检查是否有GUI模块
    try:
        import pdf_converter_gui
        has_gui = True
    except ImportError:
        has_gui = False
    
    if has_gui:
        try:
            # 集成到GUI
            print("正在整合到GUI转换流程...")
            
            # 获取原始convert_pdf函数
            original_convert_pdf = pdf_converter_gui.convert_pdf
            
            # 创建增强的convert_pdf函数
            def enhanced_convert_pdf(pdf_path, output_path, progress_callback=None, **kwargs):
                """增强的PDF转换函数，包含表格样式继承修复"""
                try:
                    # 创建转换器实例
                    converter = None
                    try:
                        from enhanced_pdf_converter import EnhancedPDFConverter
                        converter = EnhancedPDFConverter()
                    except ImportError:
                        try:
                            from improved_pdf_converter import ImprovedPDFConverter
                            converter = ImprovedPDFConverter()
                        except ImportError:
                            # 无法创建转换器，使用原始方法
                            return original_convert_pdf(pdf_path, output_path, progress_callback, **kwargs)
                    
                    # 应用表格样式修复
                    from table_style_inheritance_fix import apply_table_style_fixes
                    apply_table_style_fixes(converter)
                    
                    # 转换PDF
                    result = converter.convert_pdf_to_docx(pdf_path, output_path, 
                                                           progress_callback=progress_callback, **kwargs)
                    return result
                
                except Exception as e:
                    print(f"增强转换出错: {e}")
                    # 使用原始方法作为备用
                    return original_convert_pdf(pdf_path, output_path, progress_callback, **kwargs)
            
            # 替换GUI模块中的convert_pdf函数
            pdf_converter_gui.convert_pdf = enhanced_convert_pdf
            print("已成功整合到GUI转换流程")
            
        except Exception as e:
            print(f"整合到GUI时出错: {e}")

def find_test_pdf_files():
    """查找测试用的PDF文件"""
    test_files = []
    
    # 常见的测试目录
    test_dirs = [
        ".",
        "./test_files",
        "./samples",
        "../test_files",
        "../samples"
    ]
    
    # 在每个目录中查找PDF文件
    for test_dir in test_dirs:
        if os.path.exists(test_dir) and os.path.isdir(test_dir):
            for file in os.listdir(test_dir):
                if file.lower().endswith(".pdf"):
                    full_path = os.path.abspath(os.path.join(test_dir, file))
                    test_files.append(full_path)
    
    return test_files

def test_table_style_fix(pdf_file, output_file):
    """测试表格样式继承修复"""
    print(f"测试表格样式继承修复: {pdf_file} -> {output_file}")
    
    if not os.path.exists(pdf_file):
        print(f"错误: 输入文件不存在: {pdf_file}")
        return False
    
    # 创建转换器并应用修复
    try:
        # 导入表格样式修复模块
        from table_style_inheritance_fix import apply_table_style_fixes
        
        # 创建转换器实例
        converter = None
        try:
            from enhanced_pdf_converter import EnhancedPDFConverter
            converter = EnhancedPDFConverter()
        except ImportError:
            try:
                from improved_pdf_converter import ImprovedPDFConverter
                converter = ImprovedPDFConverter()
            except ImportError:
                print("错误: 无法创建PDF转换器实例")
                return False
        
        # 应用表格样式修复
        if not apply_table_style_fixes(converter):
            print("错误: 应用表格样式修复失败")
            return False
        
        # 转换PDF
        print("正在转换PDF...")
        result = converter.convert_pdf_to_docx(pdf_file, output_file)
        
        if os.path.exists(output_file):
            print(f"转换成功! 输出文件: {output_file}")
            return True
        else:
            print("错误: 转换失败，未生成输出文件")
            return False
        
    except Exception as e:
        print(f"测试时出错: {e}")
        traceback.print_exc()
        return False

if __name__ == "__main__":
    try:
        # 处理命令行参数
        if len(sys.argv) > 1:
            if sys.argv[1] == "test" and len(sys.argv) >= 4:
                # 测试模式
                pdf_file = sys.argv[2]
                output_file = sys.argv[3]
                test_table_style_fix(pdf_file, output_file)
            else:
                # 帮助信息
                print("用法:")
                print("  python apply_table_style_fix.py          # 应用修复")
                print("  python apply_table_style_fix.py test input.pdf output.docx  # 测试修复")
        else:
            # 默认：应用修复
            apply_table_style_fix()
    
    except Exception as e:
        print(f"执行时出错: {e}")
        traceback.print_exc()
