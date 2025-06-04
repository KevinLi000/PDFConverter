"""
直接修复 PDF Converter GUI 中的值解包错误
"""

def add_handling_for_multi_value_tuples():
    # 打开PDF转换器GUI文件
    file_path = 'pdf_converter_gui.py'
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # 创建备份
    with open(file_path + '.bak', 'w', encoding='utf-8') as f:
        f.writelines(lines)
    
    # 寻找_check_conversion_result方法中的目标位置
    target_line_index = None
    for i, line in enumerate(lines):
        if 'if isinstance(result, tuple) and len(result) == 2:' in line:
            target_line_index = i
            break
    
    if target_line_index is not None:
        # 在"if isinstance(result, tuple) and len(result) == 2:"之后插入新代码
        indent = ' ' * (len(lines[target_line_index]) - len(lines[target_line_index].lstrip()))
        new_lines = [
            f"{indent}elif isinstance(result, tuple) and len(result) > 2:\n",
            f"{indent}    # 如果元组有超过2个元素，取前两个元素\n",
            f"{indent}    status, message = result[0], result[1]\n"
        ]
        
        # 找到"else:"行
        else_index = None
        for i in range(target_line_index + 1, len(lines)):
            if 'else:' in lines[i] and '# 如果不是预期的格式' in lines[i]:
                else_index = i
                break
        
        if else_index is not None:
            # 在else行之前插入新代码
            lines[else_index:else_index] = new_lines
            
            # 写回文件
            with open(file_path, 'w', encoding='utf-8') as f:
                f.writelines(lines)
            
            print(f"成功修复了 {file_path} 中的元组解包错误！")
            print(f"已在第 {else_index} 行前添加了处理多值元组的代码。")
            print(f"原始文件已备份为 {file_path}.bak")
            return True
    
    print("未能找到目标位置，请手动检查文件。")
    return False

if __name__ == "__main__":
    add_handling_for_multi_value_tuples()
