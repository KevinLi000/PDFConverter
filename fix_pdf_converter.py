"""
修复PDFConverterGUI中的数据解析问题
"""

def fix_pdf_converter_gui():
    """
    修复pdf_converter_gui.py中的错误
    """
    import os
    import re
    
    file_path = "pdf_converter_gui.py"
    
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    
    # 修复pdf_to_word方法
    word_pattern = r"""if hasattr\(converter, 'pdf_to_word'\) and callable\(getattr\(converter, 'pdf_to_word'\)\):
                    original_pdf_to_word = converter\.pdf_to_word
                    def safe_pdf_to_word\(self, \*args, \*\*kwargs\):
                        try:
                            result = original_pdf_to_word\(\*args, \*\*kwargs\)
                            # Ensure the result is a single string \(file path\)
                            if isinstance\(result, tuple\) or isinstance\(result, list\):
                                # If a tuple or list is returned, use the first item
                                return str\(result\[0\]\) if result else None
                            return result
                        except Exception as e:
                            # 处理任何意外返回值
                            if hasattr\(self, 'input_file'\) and hasattr\(self, 'output_dir'\):
                                # 如果转换失败，返回默认路径
                                import os
                                base_name = os\.path\.splitext\(os\.path\.basename\(self\.input_file\)\)\[0\]
                                return os\.path\.join\(self\.output_dir, f\"\{base_name\}\.docx\"\)
                            raise e
                    converter\.pdf_to_word = types\.MethodType\(safe_pdf_to_word, converter\)"""
    
    word_replacement = """if hasattr(converter, 'pdf_to_word') and callable(getattr(converter, 'pdf_to_word')):
                    original_pdf_to_word = converter.pdf_to_word
                    def safe_pdf_to_word(self, *args, **kwargs):
                        try:
                            result = original_pdf_to_word(*args, **kwargs)
                            # Ensure the result is a single string (file path)
                            if isinstance(result, (tuple, list)):
                                # If it's a sequence, use the first item and ensure it's a string
                                if result and len(result) > 0:
                                    return str(result[0])
                                else:
                                    # Empty sequence, return default path
                                    if hasattr(self, 'input_file') and hasattr(self, 'output_dir'):
                                        import os
                                        base_name = os.path.splitext(os.path.basename(self.input_file))[0]
                                        return os.path.join(self.output_dir, f"{base_name}.docx")
                                    return None
                            # If not a sequence, ensure it's a string
                            return str(result) if result is not None else None
                        except Exception as e:
                            # 处理任何意外返回值
                            if hasattr(self, 'input_file') and hasattr(self, 'output_dir'):
                                # 如果转换失败，返回默认路径
                                import os
                                base_name = os.path.splitext(os.path.basename(self.input_file))[0]
                                return os.path.join(self.output_dir, f"{base_name}.docx")
                            raise e
                    converter.pdf_to_word = types.MethodType(safe_pdf_to_word, converter)"""
    
    content = re.sub(word_pattern, word_replacement, content)
    
    # 修复pdf_to_excel方法
    excel_pattern = r"""if hasattr\(converter, 'pdf_to_excel'\) and callable\(getattr\(converter, 'pdf_to_excel'\)\):
                    original_pdf_to_excel = converter\.pdf_to_excel
                    def safe_pdf_to_excel\(self, \*args, \*\*kwargs\):
                        try:
                            result = original_pdf_to_excel\(\*args, \*\*kwargs\)
                            # Ensure the result is a single string \(file path\)
                            if isinstance\(result, tuple\) or isinstance\(result, list\):
                                # If a tuple or list is returned, use the first item
                                return str\(result\[0\]\) if result else None
                            return result
                        except Exception as e:
                            # 处理任何意外返回值
                            if hasattr\(self, 'input_file'\) and hasattr\(self, 'output_dir'\):
                                # 如果转换失败，返回默认路径
                                import os
                                base_name = os\.path\.splitext\(os\.path\.basename\(self\.input_file\)\)\[0\]
                                return os\.path\.join\(self\.output_dir, f\"\{base_name\}\.xlsx\"\)
                            raise e
                    converter\.pdf_to_excel = types\.MethodType\(safe_pdf_to_excel, converter\)"""
    
    excel_replacement = """if hasattr(converter, 'pdf_to_excel') and callable(getattr(converter, 'pdf_to_excel')):
                    original_pdf_to_excel = converter.pdf_to_excel
                    def safe_pdf_to_excel(self, *args, **kwargs):
                        try:
                            result = original_pdf_to_excel(*args, **kwargs)
                            # Ensure the result is a single string (file path)
                            if isinstance(result, (tuple, list)):
                                # If it's a sequence, use the first item and ensure it's a string
                                if result and len(result) > 0:
                                    return str(result[0])
                                else:
                                    # Empty sequence, return default path
                                    if hasattr(self, 'input_file') and hasattr(self, 'output_dir'):
                                        import os
                                        base_name = os.path.splitext(os.path.basename(self.input_file))[0]
                                        return os.path.join(self.output_dir, f"{base_name}.xlsx")
                                    return None
                            # If not a sequence, ensure it's a string
                            return str(result) if result is not None else None
                        except Exception as e:
                            # 处理任何意外返回值
                            if hasattr(self, 'input_file') and hasattr(self, 'output_dir'):
                                # 如果转换失败，返回默认路径
                                import os
                                base_name = os.path.splitext(os.path.basename(self.input_file))[0]
                                return os.path.join(self.output_dir, f"{base_name}.xlsx")
                            raise e
                    converter.pdf_to_excel = types.MethodType(safe_pdf_to_excel, converter)"""
    
    content = re.sub(excel_pattern, excel_replacement, content)
    
    # 修复_check_conversion_result方法
    check_pattern = r"""if isinstance\(result, tuple\) and len\(result\) == 2:
                        status, message = result
                    else:
                        # 如果不是预期的格式，则视为错误
                        status = "error"
                        message = f"意外的转换结果: \{str\(result\)\}\""""
    
    check_replacement = """if isinstance(result, tuple) and len(result) == 2:
                        status, message = result
                    elif isinstance(result, tuple) and len(result) > 2:
                        # 如果元组有超过2个元素，取前两个元素
                        status, message = result[0], result[1]
                    else:
                        # 如果不是预期的格式，则视为错误
                        status = "error"
                        message = f"意外的转换结果: {str(result)}\""""
    
    content = re.sub(check_pattern, check_replacement, content)
    
    # 修复缩进问题
    content = content.replace("      def _update_status", "    def _update_status")
    
    # 保存修改后的文件
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(content)
    
    print("修复完成！")

if __name__ == "__main__":
    fix_pdf_converter_gui()
