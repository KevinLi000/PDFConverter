"""
针对"too many values to unpack (expected 2)"错误的直接修复
"""
import os

# 定位到PDF转换器GUI文件
file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pdf_converter_gui.py')

# 读取文件内容
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 创建文件备份
backup_path = file_path + '.bak2'
with open(backup_path, 'w', encoding='utf-8') as f:
    f.write(content)

print(f"已创建备份文件: {backup_path}")

# 找到需要修改的代码段
target_text = '''                    # 确保结果是一个二元组 (status, message)
                    if isinstance(result, tuple) and len(result) == 2:
                        status, message = result
                    else:
                        # 如果不是预期的格式，则视为错误
                        status = "error"
                        message = f"意外的转换结果: {str(result)}"'''

# 替换为修复后的代码
replacement_text = '''                    # 确保结果是一个二元组 (status, message)
                    if isinstance(result, tuple) and len(result) == 2:
                        status, message = result
                    elif isinstance(result, tuple) and len(result) > 2:
                        # 如果元组有超过2个元素，取前两个元素
                        status, message = result[0], result[1]
                    else:
                        # 如果不是预期的格式，则视为错误
                        status = "error"
                        message = f"意外的转换结果: {str(result)}"'''

# 执行替换
if target_text in content:
    new_content = content.replace(target_text, replacement_text)
    
    # 写入修改后的内容
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    print(f"成功修复了 {file_path} 中的元组解包错误！")
else:
    print("无法找到需要修改的代码段。请手动检查文件。")
