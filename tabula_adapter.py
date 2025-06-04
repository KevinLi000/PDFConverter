"""
Tabula 适配器模块 - 修复不同版本的tabula-py库之间的导入和使用差异
"""

import importlib
import sys
import traceback

def get_tabula_read_pdf():
    """
    获取tabula的read_pdf函数，适应不同版本的tabula-py库
    
    返回:
        tabula.read_pdf函数
    """
    try:
        # 首先尝试从tabula-py中导入
        import tabula
        return tabula.read_pdf
    except (ImportError, AttributeError):
        # 如果导入失败，尝试其他方式
        try:
            # 尝试从tabula.io中导入
            from tabula.io import read_pdf
            return read_pdf
        except ImportError:
            # 如果也失败，可能需要安装tabula-py
            print("错误: 无法导入tabula.read_pdf函数")
            print("请使用以下命令安装tabula-py: pip install tabula-py")
            
            # 创建一个假的函数
            def dummy_read_pdf(*args, **kwargs):
                print("tabula.read_pdf不可用，请安装tabula-py库")
                return []
            
            return dummy_read_pdf

def patch_tabula_imports():
    """
    修补tabula导入，确保可以正确导入read_pdf
    """
    try:
        # 获取read_pdf函数
        read_pdf_func = get_tabula_read_pdf()
        
        # 确保tabula模块存在
        if 'tabula' not in sys.modules:
            import tabula
        
        # 确保tabula模块有read_pdf属性
        if not hasattr(sys.modules['tabula'], 'read_pdf'):
            sys.modules['tabula'].read_pdf = read_pdf_func
            print("已成功修补tabula.read_pdf")
        
        return True
    except Exception as e:
        print(f"修补tabula导入失败: {e}")
        traceback.print_exc()
        return False

def fix_tabula_imports_in_module(module_name):
    """
    修复指定模块中的tabula导入问题
    
    参数:
        module_name: 模块名称
    """
    try:
        # 获取模块对象
        if isinstance(module_name, str):
            if module_name in sys.modules:
                module = sys.modules[module_name]
            else:
                module = importlib.import_module(module_name)
        else:
            module = module_name  # 直接传入了模块对象
        
        # 获取read_pdf函数
        read_pdf_func = get_tabula_read_pdf()
        
        # 将read_pdf函数添加到模块的全局命名空间
        module.__dict__['read_pdf'] = read_pdf_func
        
        print(f"已成功修复模块 {getattr(module, '__name__', str(module))} 中的tabula导入")
        return True
    except Exception as e:
        print(f"修复模块tabula导入失败: {e}")
        traceback.print_exc()
        return False
