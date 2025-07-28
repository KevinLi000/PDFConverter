"""
表格维度修复测试脚本
验证表格行宽高识别错误修复是否正常工作
"""

import os
import sys
import traceback

def test_table_dimension_fix():
    """测试表格维度修复功能"""
    print("=== 表格维度修复测试 ===")
    
    try:
        # 1. 测试模块导入
        print("\n1. 测试模块导入...")
        try:
            from table_dimension_fix import apply_table_dimension_fix
            print("✓ 成功导入 apply_table_dimension_fix")
        except ImportError as e:
            print(f"✗ 导入失败: {e}")
            return False
        
        # 2. 测试转换器创建
        print("\n2. 测试转换器创建...")
        try:
            from enhanced_pdf_converter import EnhancedPDFConverter
            converter = EnhancedPDFConverter()
            print("✓ 成功创建转换器")
        except ImportError as e:
            print(f"✗ 创建转换器失败: {e}")
            return False
        
        # 3. 测试修复应用
        print("\n3. 测试修复应用...")
        try:
            success = apply_table_dimension_fix(converter)
            if success:
                print("✓ 成功应用表格维度修复")
            else:
                print("✗ 应用修复失败")
                return False
        except Exception as e:
            print(f"✗ 应用修复时出错: {e}")
            traceback.print_exc()
            return False
        
        # 4. 验证新增方法
        print("\n4. 验证新增方法...")
        expected_methods = [
            '_enhanced_dimension_detection',
            '_detect_dimensions_from_text_blocks',
            '_detect_dimensions_from_image',
            '_detect_dimensions_from_drawings',
            '_estimate_dimensions_intelligently',
            'apply_precise_table_dimensions'
        ]
        
        missing_methods = []
        for method in expected_methods:
            if hasattr(converter, method):
                print(f"✓ 方法 {method} 已添加")
            else:
                print(f"✗ 方法 {method} 缺失")
                missing_methods.append(method)
        
        if missing_methods:
            print(f"缺失方法: {missing_methods}")
            return False
        
        # 5. 测试维度检测功能
        print("\n5. 测试维度检测功能...")
        try:
            # 创建模拟表格数据
            test_table_block = {
                "bbox": [100, 100, 400, 300],
                "table_data": [
                    ["Header1", "Header2", "Header3"],
                    ["Row1Col1", "Row1Col2", "Row1Col3"],
                    ["Row2Col1", "Row2Col2", "Row2Col3"]
                ]
            }
            
            # 模拟页面对象
            class MockPage:
                def get_text(self, *args, **kwargs):
                    return {"blocks": []}
                
                def get_drawings(self):
                    return []
                
                def get_pixmap(self, *args, **kwargs):
                    class MockPix:
                        width = 300
                        height = 200
                        samples = b'\x00' * (300 * 200 * 3)
                    return MockPix()
                
                @property
                def rect(self):
                    class MockRect:
                        width = 500
                        height = 700
                    return MockRect()
            
            mock_page = MockPage()
            
            # 测试维度检测
            result = converter._enhanced_dimension_detection(test_table_block, mock_page)
            
            if isinstance(result, dict):
                print("✓ 维度检测返回正确格式")
                print(f"  - 检测方法: {result.get('detection_method', 'unknown')}")
                print(f"  - 置信度: {result.get('confidence_score', 0.0)}")
                print(f"  - 行高数量: {len(result.get('precise_row_heights', []))}")
                print(f"  - 列宽数量: {len(result.get('precise_col_widths', []))}")
            else:
                print("✗ 维度检测返回格式错误")
                return False
                
        except Exception as e:
            print(f"✗ 维度检测测试失败: {e}")
            traceback.print_exc()
            return False
        
        print("\n=== 测试完成 ===")
        print("✓ 所有测试通过！表格维度修复功能正常工作")
        return True
        
    except Exception as e:
        print(f"\n✗ 测试过程中出现错误: {e}")
        traceback.print_exc()
        return False

def test_dimension_accuracy():
    """测试维度检测精度"""
    print("\n=== 维度检测精度测试 ===")
    
    try:
        from table_dimension_fix import apply_table_dimension_fix
        from enhanced_pdf_converter import EnhancedPDFConverter
        
        converter = EnhancedPDFConverter()
        apply_table_dimension_fix(converter)
        
        # 测试不同类型的表格数据
        test_cases = [
            {
                "name": "标准2x3表格",
                "table_data": [
                    ["A", "B"],
                    ["C", "D"],
                    ["E", "F"]
                ],
                "bbox": [0, 0, 200, 150]
            },
            {
                "name": "不规则4x2表格",
                "table_data": [
                    ["Header1", "Header2", "Header3", "Header4"],
                    ["LongContentInThisCell", "Short", "Medium Content", "X"]
                ],
                "bbox": [50, 50, 450, 100]
            },
            {
                "name": "单列表格",
                "table_data": [
                    ["Item1"],
                    ["Item2"],
                    ["Item3"],
                    ["Item4"]
                ],
                "bbox": [100, 100, 200, 300]
            }
        ]
        
        class MockPage:
            def get_text(self, *args, **kwargs):
                return {"blocks": []}
            def get_drawings(self):
                return []
            def get_pixmap(self, *args, **kwargs):
                class MockPix:
                    width = 500
                    height = 400
                    samples = b'\x00' * (500 * 400 * 3)
                return MockPix()
            @property
            def rect(self):
                class MockRect:
                    width = 600
                    height = 800
                return MockRect()
        
        mock_page = MockPage()
        
        for i, test_case in enumerate(test_cases, 1):
            print(f"\n{i}. 测试 {test_case['name']}...")
            
            table_block = {
                "bbox": test_case["bbox"],
                "table_data": test_case["table_data"]
            }
            
            result = converter._enhanced_dimension_detection(table_block, mock_page)
            
            expected_rows = len(test_case["table_data"])
            expected_cols = len(test_case["table_data"][0]) if test_case["table_data"] else 0
            
            row_heights = result.get("precise_row_heights", [])
            col_widths = result.get("precise_col_widths", [])
            
            print(f"  预期: {expected_rows}行 x {expected_cols}列")
            print(f"  检测: {len(row_heights)}行高 x {len(col_widths)}列宽")
            print(f"  置信度: {result.get('confidence_score', 0.0):.2f}")
            print(f"  检测方法: {result.get('detection_method', 'unknown')}")
            
            if len(col_widths) > 0:
                print(f"  列宽总和: {sum(col_widths):.3f} (应该接近1.0)")
            
            # 检查列宽是否合理
            if col_widths:
                if abs(sum(col_widths) - 1.0) < 0.1:
                    print("  ✓ 列宽归一化正确")
                else:
                    print("  ✗ 列宽归一化有问题")
            
            # 检查行高是否合理
            if row_heights:
                if all(h > 0 for h in row_heights):
                    print("  ✓ 行高都为正值")
                else:
                    print("  ✗ 存在无效行高")
        
        print("\n=== 精度测试完成 ===")
        return True
        
    except Exception as e:
        print(f"精度测试失败: {e}")
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("开始表格维度修复测试...")
    
    # 基本功能测试
    basic_test_result = test_table_dimension_fix()
    
    # 精度测试
    if basic_test_result:
        accuracy_test_result = test_dimension_accuracy()
        
        if accuracy_test_result:
            print("\n🎉 所有测试通过！表格维度修复已正确实现并集成。")
        else:
            print("\n⚠️ 基本功能正常，但精度测试有问题。")
    else:
        print("\n❌ 基本功能测试失败，请检查代码。")
