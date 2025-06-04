#!/usr/bin/env python
"""
应用全面的PDF转换器修复
这个脚本集成并应用所有的修复，解决:
1. 'Page' object has no attribute 'find_tables'错误
2. 'dict' object has no attribute 'cells'错误
3. 表格处理和提取中的各种兼容性问题
4. tabula导入和使用问题 - 'can't import name read_pdf from tabula'
5. 方法名称适配问题 - 'EnhancedPDFConverter' not being associated with a value
"""

import os
import sys
import traceback
import importlib.util

def check_module_exists(module_name):
    """检查模块是否已安装"""
    try:
        importlib.import_module(module_name)
        return True
    except ImportError:
        return False

def main():
    """应用所有PDF转换器的修复和增强"""
    print("PDF转换器全面修复程序")
    print("=" * 50)
      # 检查必要的模块是否存在
    required_modules = [
        "fitz",        # PyMuPDF
        "docx"         # python-docx
    ]
    
    recommended_modules = [
        "cv2",         # OpenCV
        "numpy",       # NumPy
        "PIL",         # Pillow
        "tabula"       # tabula-py
    ]
    
    missing_required = []
    missing_recommended = []
    
    for module in required_modules:
        if not check_module_exists(module):
            missing_required.append(module)
    
    for module in recommended_modules:
        if not check_module_exists(module):
            missing_recommended.append(module)
    
    if missing_required:
        print("错误: 缺少以下必要模块:")
        for module in missing_required:
            print(f" - {module}")
        print("\n请使用以下命令安装所需模块:")
          pip_modules = {
            "fitz": "PyMuPDF",
            "docx": "python-docx"
        }
        
        install_cmd = "pip install " + " ".join([pip_modules.get(m, m) for m in missing_required])
        print(install_cmd)
        return 1
    
    if missing_recommended:
        print("警告: 缺少以下推荐模块 (一些高级功能可能不可用):")
        for module in missing_recommended:
            print(f" - {module}")
          pip_modules = {
            "cv2": "opencv-python",
            "PIL": "Pillow",
            "tabula": "tabula-py"
        }
        
        install_cmd = "pip install " + " ".join([pip_modules.get(m, m) for m in missing_recommended])
        print(f"推荐安装: {install_cmd}")
    
    # 应用全面修复
    fixes_applied = []
    
    try:
        print("\n应用全面PDF转换器修复...")
        
        # 1. 导入增强的PDF转换器
        try:
            from enhanced_pdf_converter import EnhancedPDFConverter
            
            # 2. 创建临时实例
            converter_instance = EnhancedPDFConverter()
              # 3. 应用全面修复
            import comprehensive_pdf_fix
            comprehensive_pdf_fix.apply_comprehensive_fixes(converter_instance)
              # 4. 特别应用tabula导入修复
            try:
                import tabula_adapter
                tabula_adapter.patch_tabula_imports()
                fixes_applied.append("tabula导入修复")
                print("已成功应用tabula导入修复")
            except ImportError:
                print("警告: 无法导入tabula_adapter模块")
                
            # 5. 应用方法名称适配
            try:
                import method_name_adapter
                method_name_adapter.apply_method_name_adaptations(converter_instance)
                fixes_applied.append("方法名称适配")
                print("已成功应用方法名称适配")
            except ImportError:
                print("警告: 无法导入method_name_adapter模块")
            
            fixes_applied.append("全面PDF转换器修复")
            print("已成功应用全面PDF转换器修复")
        except ImportError as e:
            print(f"错误: 无法导入EnhancedPDFConverter: {e}")
            print("请确保enhanced_pdf_converter.py文件存在于当前目录")
            return 1
        except Exception as e:
            print(f"错误: 应用全面修复时失败: {e}")
            traceback.print_exc()
            
            # 尝试应用个别修复
            try:
                print("\n尝试应用单独的修复...")
                
                # 1. 表格检测修复
                try:
                    import fix_table_detection
                    fix_table_detection.apply_table_detection_patch()
                    fixes_applied.append("表格检测修复")
                    print("已应用表格检测修复")
                except Exception as e:
                    print(f"表格检测修复应用失败: {e}")
                
                # 2. Dict cells错误修复
                try:
                    import fix_dict_cells_error
                    fix_dict_cells_error.apply_dict_cells_fix(converter_instance)
                    fixes_applied.append("Dict cells错误修复")
                    print("已应用Dict cells错误修复")
                except Exception as e:
                    print(f"Dict cells错误修复应用失败: {e}")
                
                # 3. PDF转换器修复
                try:
                    import pdf_converter_fix
                    pdf_converter_fix.apply_enhanced_pdf_converter_fixes(converter_instance)
                    fixes_applied.append("PDF转换器修复")
                    print("已应用PDF转换器修复")
                except Exception as e:
                    print(f"PDF转换器修复应用失败: {e}")
                
                # 4. 直接表格检测补丁
                try:
                    import direct_table_detection_patch
                    if hasattr(direct_table_detection_patch, "apply_patch"):
                        direct_table_detection_patch.apply_patch()
                        fixes_applied.append("直接表格检测补丁")
                        print("已应用直接表格检测补丁")
                except Exception as e:
                    print(f"直接表格检测补丁应用失败: {e}")
            except Exception as e:
                print(f"应用单独修复时出错: {e}")
    except Exception as e:
        print(f"应用修复时出现错误: {e}")
        traceback.print_exc()
    
    # 显示结果
    if fixes_applied:
        print("\n成功应用的修复:")
        for fix in fixes_applied:
            print(f" - {fix}")
        print("\nPDF转换器修复已完成！现在应该可以正常使用了。")
        return 0
    else:
        print("\n错误: 未能应用任何修复")
        return 1

if __name__ == "__main__":
    sys.exit(main())
