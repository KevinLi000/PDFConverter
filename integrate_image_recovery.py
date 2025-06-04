"""
将图像恢复增强集成到PDF转换流程中
"""

import os
import sys
import types
import importlib.util

def import_module_from_path(module_name, file_path):
    """从文件路径导入模块"""
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

def integrate_image_recovery():
    """集成图像恢复增强到现有的PDF转换器集成流程"""
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 导入图像恢复增强模块
        image_recovery_path = os.path.join(current_dir, "image_recovery_enhancement.py")
        if not os.path.exists(image_recovery_path):
            print("错误: 找不到图像恢复增强模块")
            return False
        
        image_recovery = import_module_from_path("image_recovery_enhancement", image_recovery_path)
        
        # 导入现有的集成器模块
        integrator_path = os.path.join(current_dir, "all_pdf_fixes_integrator.py")
        if not os.path.exists(integrator_path):
            print("错误: 找不到PDF修复集成器")
            return False
        
        integrator = import_module_from_path("all_pdf_fixes_integrator", integrator_path)
        
        # 获取原始的integrate_all_fixes函数
        if hasattr(integrator, 'integrate_all_fixes'):
            original_integrate = integrator.integrate_all_fixes
            
            # 创建增强版本的集成函数
            def enhanced_integrate_all_fixes(converter, format_preservation_level=None):
                """增强版本的集成函数，添加图像恢复增强"""
                # 首先调用原始集成函数
                result = original_integrate(converter, format_preservation_level)
                
                # 然后应用图像恢复增强
                print("\n== 应用图像恢复增强 ==")
                image_recovery.apply_image_recovery(converter)
                
                return result
            
            # 替换原始集成函数
            integrator.integrate_all_fixes = enhanced_integrate_all_fixes
            print("已成功集成图像恢复增强到PDF修复流程")
            
            # 导入并修改增强型PDF转换器
            converter_path = os.path.join(current_dir, "enhanced_pdf_converter.py")
            if os.path.exists(converter_path):
                try:
                    # 复制_temp_complex_page_method.py中的方法到enhanced_pdf_converter.py
                    complex_page_method_path = os.path.join(current_dir, "_temp_complex_page_method.py")
                    if os.path.exists(complex_page_method_path):
                        # 读取_is_complex_page方法
                        with open(complex_page_method_path, 'r', encoding='utf-8') as f:
                            complex_page_content = f.read()
                        
                        # 检查方法是否存在于转换器中
                        with open(converter_path, 'r', encoding='utf-8') as f:
                            converter_content = f.read()
                        
                        if '_is_complex_page' not in converter_content:
                            # 将方法添加到转换器类定义的末尾
                            with open(converter_path, 'r', encoding='utf-8') as f:
                                lines = f.readlines()
                            
                            # 查找类定义结束的位置
                            class_end_idx = -1
                            for i in range(len(lines)-1, -1, -1):
                                if lines[i].strip() == "# End of EnhancedPDFConverter class" or lines[i].strip() == "# 结束 EnhancedPDFConverter 类":
                                    class_end_idx = i
                                    break
                            
                            if class_end_idx == -1:
                                # 如果没有找到明确的结束标记，尝试查找类定义
                                class_defined = False
                                brace_count = 0
                                for i, line in enumerate(lines):
                                    if "class EnhancedPDFConverter" in line:
                                        class_defined = True
                                    if class_defined:
                                        if "{" in line:
                                            brace_count += 1
                                        if "}" in line:
                                            brace_count -= 1
                                            if brace_count == 0:
                                                class_end_idx = i
                                                break
                            
                            # 如果找到了插入位置
                            if class_end_idx != -1:
                                # 提取_is_complex_page和_add_border_to_picture方法
                                import re
                                is_complex_page_match = re.search(r'def _is_complex_page\(.*?return.*?\)', complex_page_content, re.DOTALL)
                                add_border_match = re.search(r'def _add_border_to_picture\(.*?return.*?\)', complex_page_content, re.DOTALL)
                                
                                methods_to_add = ""
                                
                                if is_complex_page_match:
                                    methods_to_add += "\n    " + is_complex_page_match.group(0).replace("\n", "\n    ") + "\n"
                                
                                if add_border_match:
                                    methods_to_add += "\n    " + add_border_match.group(0).replace("\n", "\n    ") + "\n"
                                
                                if methods_to_add:
                                    lines.insert(class_end_idx, methods_to_add)
                                    
                                    with open(converter_path, 'w', encoding='utf-8') as f:
                                        f.writelines(lines)
                                    
                                    print("已成功将复杂页面检测和图像边框添加方法集成到PDF转换器")
                    
                except Exception as e:
                    print(f"修改PDF转换器时出错: {e}")
                    import traceback
                    traceback.print_exc()
            
            return True
        else:
            print("错误: 集成器模块中找不到integrate_all_fixes函数")
            return False
        
    except Exception as e:
        print(f"集成图像恢复增强时出错: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = integrate_image_recovery()
    print(f"集成结果: {'成功' if success else '失败'}")
    
    if success:
        print("\n使用说明:")
        print("1. 导入 all_pdf_fixes_integrator 模块")
        print("2. 调用 integrate_all_fixes 函数应用所有修复，包括图像恢复增强")
        print("3. 所有的PDF转换将自动应用图像恢复增强，确保图像不会丢失")
