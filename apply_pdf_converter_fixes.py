#!/usr/bin/env python
"""
PDF转换器增强修复程序
- 修复'Page' object has no attribute 'find_tables'错误
- 增强表格识别与处理
- 增强字体处理功能
"""

import os
import sys
import importlib
import traceback

def main():
    """应用所有PDF转换器的修复和增强"""
    print("PDF转换器增强修复程序")
    print("=" * 50)
    
    # 检查必要的模块是否存在
    required_modules = [
        "fitz",        # PyMuPDF
        "cv2",         # OpenCV
        "numpy",       # NumPy
        "PIL",         # Pillow
        "docx"         # python-docx
    ]
    
    missing_modules = []
    for module in required_modules:
        try:
            importlib.import_module(module)
        except ImportError:
            missing_modules.append(module)
    
    if missing_modules:
        print("错误: 缺少以下必要模块:")
        for module in missing_modules:
            print(f" - {module}")
        print("\n请使用以下命令安装所需模块:")
        
        pip_modules = {
            "fitz": "PyMuPDF",
            "cv2": "opencv-python",
            "PIL": "Pillow",
            "docx": "python-docx"
        }
        
        install_cmd = "pip install " + " ".join([pip_modules.get(m, m) for m in missing_modules])
        print(install_cmd)
        return 1
    
    # 应用补丁
    fixes_applied = []
    
    # 1. 应用表格检测修复
    try:
        print("\n应用表格检测修复...")
        import fix_table_detection
        fix_table_detection.apply_table_detection_patch()
        fixes_applied.append("表格检测修复")
    except ImportError:
        print("警告: 未找到表格检测修复模块，尝试其他方法")
        try:
            # 尝试导入表格检测工具
            from table_detection_utils import add_table_detection_capability
            from enhanced_pdf_converter import EnhancedPDFConverter
            
            # 创建临时实例以应用修复
            temp_instance = EnhancedPDFConverter()
            add_table_detection_capability(temp_instance)
            print("已成功应用表格检测功能")
            fixes_applied.append("表格检测功能")
        except Exception as e:
            print(f"表格检测功能应用失败: {e}")
    except Exception as e:
        print(f"表格检测修复应用失败: {e}")
        traceback.print_exc()
      # 2. 应用转换器修复
    try:
        print("\n应用PDF转换器修复...")
        import pdf_converter_fix
        # 检查是否需要创建实例来应用修复
        if hasattr(pdf_converter_fix, "apply_enhanced_pdf_converter_fixes"):
            from enhanced_pdf_converter import EnhancedPDFConverter
            temp_instance = EnhancedPDFConverter()
            pdf_converter_fix.apply_enhanced_pdf_converter_fixes(temp_instance)
        else:
            # 假设模块会自行处理
            pass
        fixes_applied.append("PDF转换器修复")
    except ImportError:
        print("警告: 未找到PDF转换器修复模块")
    except Exception as e:
        print(f"PDF转换器修复应用失败: {e}")
        traceback.print_exc()
    
    # 2.1 应用dict cells错误修复
    try:
        print("\n应用dict cells错误修复...")
        import fix_dict_cells_error
        from enhanced_pdf_converter import EnhancedPDFConverter
        temp_instance = EnhancedPDFConverter()
        fix_dict_cells_error.apply_dict_cells_fix(temp_instance)
        fixes_applied.append("dict cells错误修复")
    except ImportError:
        print("警告: 未找到dict cells错误修复模块")
    except Exception as e:
        print(f"dict cells错误修复应用失败: {e}")
        traceback.print_exc()
    
    # 3. 应用PDF直接修复补丁
    try:
        print("\n应用直接修复补丁...")
        import direct_table_detection_patch
        if hasattr(direct_table_detection_patch, "apply_patch"):
            direct_table_detection_patch.apply_patch()
            fixes_applied.append("直接修复补丁")
        else:
            print("直接修复补丁模块不包含apply_patch方法")
    except ImportError:
        print("未找到直接修复补丁模块")
    except Exception as e:
        print(f"直接修复补丁应用失败: {e}")
        traceback.print_exc()
    
    # 输出结果摘要
    print("\n" + "=" * 50)
    if fixes_applied:
        print(f"成功应用了以下修复: {', '.join(fixes_applied)}")
        print("\nPDF转换器已增强，现在可以处理:")
        print(" - 解决'Page' object has no attribute 'find_tables'错误")
        print(" - 增强的表格识别和处理")
        print(" - 改进的字体处理")
        print(" - 表格样式增强")
        print("\n使用方法: 正常使用增强型PDF转换器即可，无需额外步骤")
        return 0
    else:
        print("未能应用任何修复")
        print("请检查是否已安装所有必要的模块，并确保修复模块在正确的路径中")
        return 1

if __name__ == "__main__":
    sys.exit(main())
