"""
将所有PDF转换修复应用到GUI
此脚本用于将所有PDF转换修复应用到GUI程序
"""

import os
import sys
import traceback

def main():
    """主函数"""
    print("=" * 50)
    print("应用所有PDF转换修复到GUI")
    print("=" * 50)
    
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 添加当前目录到路径
    if current_dir not in sys.path:
        sys.path.insert(0, current_dir)
    
    # 导入GUI模块
    try:
        import pdf_converter_gui
        print("✓ 成功导入PDF转换器GUI模块")
    except ImportError:
        print("✗ 导入PDF转换器GUI模块失败")
        print("请确保pdf_converter_gui.py文件存在")
        return
    
    # 导入集成修复模块
    try:
        import all_pdf_fixes_integrator
        print("✓ 成功导入集成修复模块")
    except ImportError:
        print("✗ 导入集成修复模块失败")
        print("请确保all_pdf_fixes_integrator.py文件存在")
        return
    
    # 应用修复到GUI类
    try:
        # 检查GUI类是否存在
        if hasattr(pdf_converter_gui, 'PDFConverterGUI'):
            # 获取原始的on_convert_button_click方法
            original_on_convert = pdf_converter_gui.PDFConverterGUI.on_convert_button_click
            
            # 定义增强版方法
            def enhanced_on_convert(self, event=None):
                """增强版的转换按钮处理方法，应用所有修复"""
                print("正在应用所有PDF转换修复...")
                
                # 获取转换器实例
                if hasattr(self, 'converter'):
                    # 应用所有修复
                    self.converter = all_pdf_fixes_integrator.integrate_all_fixes(self.converter)
                    print("已成功应用所有修复到转换器")
                
                # 调用原始方法
                return original_on_convert(self, event)
            
            # 替换原始方法
            pdf_converter_gui.PDFConverterGUI.on_convert_button_click = enhanced_on_convert
            print("✓ 已成功修改GUI类，将在转换时应用所有修复")
        else:
            print("✗ 未找到PDFConverterGUI类")
            return
    except Exception as e:
        print(f"✗ 应用修复到GUI类时出错: {e}")
        traceback.print_exc()
        return
    
    # 尝试运行GUI
    try:
        print("\n正在启动修复后的PDF转换器GUI...")
        
        if hasattr(pdf_converter_gui, 'main'):
            pdf_converter_gui.main()
        else:
            # 如果没有main函数，创建GUI实例并启动
            app = pdf_converter_gui.tk.Tk()
            app.title("增强型PDF转换工具 (已修复)")
            gui = pdf_converter_gui.PDFConverterGUI(app)
            app.mainloop()
    except Exception as e:
        print(f"✗ 启动GUI时出错: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    main()
