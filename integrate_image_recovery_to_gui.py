"""
将图像恢复增强集成到PDF转换器GUI中
"""

import os
import sys
import importlib.util
import time

def import_module_from_path(module_name, file_path):
    """从文件路径导入模块"""
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

def integrate_image_recovery_to_gui():
    """将图像恢复增强集成到PDF转换器GUI中"""
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 检查GUI文件是否存在
        gui_path = os.path.join(current_dir, "pdf_converter_gui.py")
        if not os.path.exists(gui_path):
            print("错误: 找不到pdf_converter_gui.py文件")
            return False
        
        # 读取GUI文件内容
        with open(gui_path, "r", encoding="utf-8") as f:
            gui_content = f.read()
        
        # 检查是否已经集成了图像恢复增强
        if "image_recovery_enhancement" in gui_content:
            print("图像恢复增强已经集成到GUI中")
            return True
        
        # 找到导入语句部分
        import_section_end = gui_content.find("class PDFConverterGUI")
        if import_section_end == -1:
            print("错误: 无法找到GUI类定义")
            return False
        
        # 找到最后一个导入语句
        last_import_idx = gui_content.rfind("import", 0, import_section_end)
        if last_import_idx == -1:
            print("错误: 无法找到导入语句")
            return False
        
        # 找到此行的结束位置
        line_end = gui_content.find("\n", last_import_idx)
        if line_end == -1:
            line_end = import_section_end
        
        # 插入图像恢复增强的导入语句
        new_imports = """
# 图像恢复增强模块
try:
    from image_recovery_enhancement import enhance_image_extraction
    has_image_recovery = True
except ImportError:
    has_image_recovery = False
"""
        updated_content = gui_content[:line_end+1] + new_imports + gui_content[line_end+1:]
        
        # 找到convert_pdf方法
        convert_method_start = updated_content.find("def convert_pdf(self")
        if convert_method_start == -1:
            print("错误: 无法找到convert_pdf方法")
            return False
        
        # 找到方法体开始
        method_body_start = updated_content.find(":", convert_method_start)
        if method_body_start == -1:
            print("错误: 无法找到convert_pdf方法体")
            return False
        
        # 找到第一行非注释代码
        next_line = updated_content.find("\n", method_body_start)
        if next_line == -1:
            print("错误: 无法找到convert_pdf方法的第一行代码")
            return False
        
        # 找到第一个非空白字符
        code_start = next_line + 1
        while code_start < len(updated_content) and (updated_content[code_start].isspace() or updated_content[code_start] == '#'):
            if updated_content[code_start] == '#':
                code_start = updated_content.find("\n", code_start) + 1
            else:
                code_start += 1
        
        # 插入图像恢复增强代码
        image_recovery_code = """        # 应用图像恢复增强
        if has_image_recovery and self.format_preservation_var.get() == "maximum":
            try:
                enhance_image_extraction(self.converter)
                self.update_status("已应用图像恢复增强...")
            except Exception as e:
                print(f"应用图像恢复增强时出错: {e}")
        
"""
        updated_content = updated_content[:code_start] + image_recovery_code + updated_content[code_start:]
        
        # 添加GUI选项说明
        option_idx = updated_content.find("self.format_preservation_var.set(")
        if option_idx != -1:
            # 找到这行的结束位置
            option_line_end = updated_content.find("\n", option_idx)
            if option_line_end != -1:
                # 找到RadioButton部分
                radio_section = updated_content.find("Radiobutton(format_frame", option_line_end)
                if radio_section != -1:
                    # 获取当前缩进级别
                    indent = ""
                    i = radio_section - 1
                    while i >= 0 and updated_content[i].isspace():
                        indent = updated_content[i] + indent
                        i -= 1
                    
                    # 找到最后一个RadioButton
                    last_radio = updated_content.rfind("Radiobutton(format_frame", 0, radio_section + 500)
                    if last_radio != -1:
                        last_radio_end = updated_content.find("\n", last_radio)
                        if last_radio_end != -1:
                            # 添加提示信息
                            hint_text = f"\n{indent}# 添加图像恢复提示\n{indent}ttk.Label(format_frame, text=\"注: '最大化保留'模式会启用图像恢复增强\", font=(None, 8, 'italic')).pack(anchor='w', padx=20, pady=(0, 5))\n"
                            updated_content = updated_content[:last_radio_end+1] + hint_text + updated_content[last_radio_end+1:]
        
        # 保存更新后的文件
        # 先备份原文件
        backup_path = os.path.join(current_dir, f"pdf_converter_gui.py.bak.{int(time.time())}")
        with open(backup_path, "w", encoding="utf-8") as f:
            f.write(gui_content)
        print(f"已创建原文件备份: {backup_path}")
        
        # 保存更新后的文件
        with open(gui_path, "w", encoding="utf-8") as f:
            f.write(updated_content)
        
        print("已成功将图像恢复增强集成到GUI中")
        print("使用'最大化保留'模式时会自动启用图像恢复增强")
        
        return True
    except Exception as e:
        print(f"集成图像恢复增强到GUI时出错: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = integrate_image_recovery_to_gui()
    if success:
        print("\n集成完成，现在可以运行GUI应用:")
        print("python pdf_converter_gui.py")
    else:
        print("\n集成失败，请检查错误信息")
