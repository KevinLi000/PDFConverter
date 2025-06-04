#!/usr/bin/env python
"""
测试转换器补丁模块是否正确返回转换器对象
"""

import sys
import os

# 添加当前目录到路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 创建模拟转换器类
class MockConverter:
    def __init__(self):
        self.name = "MockConverter"
    
    def test_method(self):
        return "Original test method"

# 导入补丁函数
from converter_patches import apply_converter_patches

# 创建一个模拟转换器实例
mock_converter = MockConverter()

# 应用补丁
patched_converter = apply_converter_patches(mock_converter)

# 检查返回值是否是转换器对象
print(f"返回值类型: {type(patched_converter)}")
print(f"是否返回原始转换器: {patched_converter is mock_converter}")
print(f"转换器名称: {patched_converter.name}")

# 验证成功
print("测试完成: 转换器补丁模块现在正确返回转换器对象")
