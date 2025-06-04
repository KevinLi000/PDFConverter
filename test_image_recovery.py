"""
测试图像恢复增强模块
"""

import os
import sys
import importlib.util

def import_module_from_path(module_name, file_path):
    """从文件路径导入模块"""
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

def main():
    try:
        # 导入所需的模块
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 1. 导入图像恢复增强模块
        image_recovery_path = os.path.join(current_dir, "image_recovery_enhancement.py")
        if os.path.exists(image_recovery_path):
            image_recovery = import_module_from_path("image_recovery_enhancement", image_recovery_path)
            print("成功导入图像恢复增强模块")
        else:
            print("错误: 找不到图像恢复增强模块")
            return False
        
        # 2. 导入PDF转换器
        converter_path = os.path.join(current_dir, "enhanced_pdf_converter.py")
        if os.path.exists(converter_path):
            converter_module = import_module_from_path("enhanced_pdf_converter", converter_path)
            if hasattr(converter_module, "EnhancedPDFConverter"):
                converter_class = converter_module.EnhancedPDFConverter
                print("成功导入PDF转换器")
            else:
                print("错误: 找不到EnhancedPDFConverter类")
                return False
        else:
            print("错误: 找不到enhanced_pdf_converter.py")
            return False
        
        # 3. 检查是否可以创建转换器实例
        try:
            temp_dir = os.path.join(current_dir, "temp")
            os.makedirs(temp_dir, exist_ok=True)
            
            # 创建转换器实例
            converter = converter_class()
            print("成功创建转换器实例")
            
            # 应用图像恢复增强
            image_recovery.apply_image_recovery(converter)
            print("已应用图像恢复增强")
            
            return True
        except Exception as e:
            print(f"创建转换器实例或应用增强时出错: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    except Exception as e:
        print(f"测试过程中出错: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = main()
    print(f"测试结果: {'成功' if success else '失败'}")
