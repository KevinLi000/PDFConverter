"""
简化的PDF转换器修复工具
修复"too many values to unpack (expected 2)"错误
"""

# 复制这个小脚本，用来修复pdf_converter_gui.py文件中的解包错误
# 使用方法：复制这个脚本，然后在命令行运行：python fix_unpacking_error.py

import os
import re

def fix_pdf_converter_gui():
    """修复pdf_converter_gui.py中的元组解包错误"""
    
    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pdf_converter_gui.py')
    
    if not os.path.exists(file_path):
        print(f"错误：找不到文件 {file_path}")
        return False
    
    # 读取文件内容
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    
    # 创建备份
    backup_path = file_path + '.bak'
    with open(backup_path, 'w', encoding='utf-8') as file:
        file.write(content)
    
    # 修复缩进问题
    content = content.replace("      def _update_status", "    def _update_status")
    
    # 在_check_conversion_result方法中修复解包错误
    pattern = r"(if isinstance\(result, tuple\) and len\(result\) == 2:\s+status, message = result\s+)(else:\s+# 如果不是预期的格式，则视为错误\s+status = \"error\"\s+message = f\"意外的转换结果: \{str\(result\)\}\")"
    
    replacement = r"\1elif isinstance(result, tuple) and len(result) > 2:\n                        # 如果元组有超过2个元素，取前两个元素\n                        status, message = result[0], result[1]\n                    \2"
    
    modified_content = re.sub(pattern, replacement, content)
    
    # 修复原始的pdf_to_word和pdf_to_excel方法
    # 修复pdf_to_word方法
    word_pattern = r"(if isinstance\(result, tuple\) or isinstance\(result, list\):\s+# If a tuple or list is returned, use the first item\s+return str\(result\[0\]\) if result else None\s+return result)"
    
    word_replacement = r"if isinstance(result, (tuple, list)):\n                                # If it's a sequence, use the first item and ensure it's a string\n                                if result and len(result) > 0:\n                                    return str(result[0])\n                                else:\n                                    # Empty sequence, return default path\n                                    if hasattr(self, 'input_file') and hasattr(self, 'output_dir'):\n                                        import os\n                                        base_name = os.path.splitext(os.path.basename(self.input_file))[0]\n                                        return os.path.join(self.output_dir, f\"{base_name}.docx\")\n                                    return None\n                            # If not a sequence, ensure it's a string\n                            return str(result) if result is not None else None"
    
    modified_content = re.sub(word_pattern, word_replacement, modified_content)
    
    # 修复pdf_to_excel方法
    excel_pattern = r"(if isinstance\(result, tuple\) or isinstance\(result, list\):\s+# If a tuple or list is returned, use the first item\s+return str\(result\[0\]\) if result else None\s+return result)"
    
    excel_replacement = r"if isinstance(result, (tuple, list)):\n                                # If it's a sequence, use the first item and ensure it's a string\n                                if result and len(result) > 0:\n                                    return str(result[0])\n                                else:\n                                    # Empty sequence, return default path\n                                    if hasattr(self, 'input_file') and hasattr(self, 'output_dir'):\n                                        import os\n                                        base_name = os.path.splitext(os.path.basename(self.input_file))[0]\n                                        return os.path.join(self.output_dir, f\"{base_name}.xlsx\")\n                                    return None\n                            # If not a sequence, ensure it's a string\n                            return str(result) if result is not None else None"
    
    modified_content = re.sub(excel_pattern, excel_replacement, modified_content)
    
    # 保存修改后的内容
    if content != modified_content:
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(modified_content)
        print(f"成功修复了 {file_path} 文件中的解包错误！")
        print(f"原始文件已备份为 {backup_path}")
        return True
    else:
        print("未发现需要修复的问题。")
        return False

if __name__ == "__main__":
    fix_pdf_converter_gui()
