"""
PDF转换器全面修复使用指南
=======================

这个文档解释如何使用comprehensive_pdf_fix.py和其他修复模块来解决PDF转换器中的问题。

基本问题修复
----------

问题1: 'Page' object has no attribute 'find_tables'
------------------------------------------
这个错误发生在使用较旧版本的PyMuPDF时，因为find_tables方法是在较新版本中添加的。

问题2: 'dict' object has no attribute 'cells'
------------------------------------------
这个错误发生在处理表格对象时，因为某些表格可能是以字典形式表示的，而不是具有cells属性的对象。

新增增强修复
----------

问题3: 表格样式和边框问题
------------------------------------------
表格在转换后可能缺少边框或边框不清晰，影响表格的可读性和外观。

问题4: 图像识别失败
------------------------------------------
某些图像在转换过程中无法被正确识别和提取，导致图像缺失。

问题5: 图像尺寸和位置问题
------------------------------------------
即使图像被成功识别，其尺寸和位置可能不正确，导致布局问题。

使用方法
-------

1. 确保已安装必要的依赖:
   ```
   pip install PyMuPDF python-docx pillow opencv-python numpy
   ```

2. 使用全面集成修复 (推荐):

   ```python
   from enhanced_pdf_converter import EnhancedPDFConverter
   import all_pdf_fixes_integrator

   # 创建转换器实例
   converter = EnhancedPDFConverter()
   
   # 应用所有修复（包括表格样式和图像处理修复）
   converter = all_pdf_fixes_integrator.integrate_all_fixes(converter)
   
   # 使用修复后的转换器
   converter.convert_pdf_to_docx("input.pdf", "output.docx")
   ```

3. 使用内置的GUI启动脚本:
   ```
   python run_gui_with_all_fixes.py
   ```
   这将启动已应用所有修复的PDF转换器GUI界面。

4. 单独应用特定修复:   ```python
   from enhanced_pdf_converter import EnhancedPDFConverter
   
   # 应用表格样式修复
   import table_detection_style_fix
   converter = EnhancedPDFConverter()
   converter = table_detection_style_fix.fix_table_detection_and_style(converter)
   
   # 应用图像处理修复
   import table_image_fix
   converter = table_image_fix.apply_image_fixes(converter)
   
   # 应用原始表格检测修复
   import comprehensive_pdf_fix
   converter = comprehensive_pdf_fix.apply_comprehensive_fixes(converter)
   ```

修复详情
-------

### 1. 表格样式和边框修复

表格样式修复主要解决以下问题:
- 增强表格边框粗细，从4pt增加到8pt，使边框更加明显
- 为每个单元格添加显式边框设置，确保所有单元格边框都能正确显示
- 改进表格样式的继承机制，确保样式能够正确应用

主要修改文件:
- `table_detection_style_fix.py`: 专用于修复表格样式和边框问题
- `enhanced_pdf_converter.py`: 在`_process_table_block`方法中增强了边框设置

### 2. 图像识别和处理修复

图像处理修复解决以下问题:
- 实现多种图像提取方法，提高图像识别的成功率
  - 通过xref提取嵌入图片
  - 使用边界框(bbox)直接裁剪页面区域
  - 使用扩展边界框捕获可能被错误裁剪的图像
- 增加图像提取分辨率，从2.0提高到4.0，大幅提升图像质量
- 自动选择最佳图像质量的提取结果

主要修改文件:
- `table_image_fix.py`: 专用于修复图像处理问题
- `enhanced_pdf_converter.py`: 在`_process_image_block_enhanced`方法中增强了图像提取

### 3. 全面集成

为了确保所有修复无缝协作，我们创建了集成修复模块:
- `all_pdf_fixes_integrator.py`: 整合所有修复，确保它们按正确顺序应用
- `run_gui_with_all_fixes.py`: 提供一键启动应用所有修复的GUI界面

测试修复效果
----------

可以使用测试脚本验证修复是否正确应用:
```
python test_all_pdf_fixes.py
```

这个脚本会检查所有修复是否正确应用，包括:
- 表格边框增强
- 图像处理增强
- 表格区域标记整合

已知限制
-------

尽管进行了全面修复，但仍存在以下限制:
1. 极为复杂的表格格式可能无法完全保留
2. 某些特殊的PDF图像格式可能仍然无法正确识别
3. 表格内的特殊格式（如单元格内的小图像）可能会丢失

更新日志
-------

### 2025-05-29
- 增加表格边框增强修复
- 增加多方法图像提取和高分辨率支持
- 创建全面集成修复模块
- 添加图像尺寸和位置修复
- 改进表格区域标记整合
"""
   converter = EnhancedPDFConverter()
   
   # 应用表格检测修复
   fix_table_detection.apply_table_detection_patch()
   
   # 应用dict cells错误修复
   converter = fix_dict_cells_error.apply_dict_cells_fix(converter)
   
   # 现在可以正常使用converter
   converter.convert_pdf_to_docx("input.pdf", "output.docx")
   ```

修复原理
-------

1. 表格检测修复:
   - 提供备用的表格检测方法，在PyMuPDF的find_tables不可用时使用
   - 使用页面的线条和文本块来识别表格结构

2. Dict cells错误修复:
   - 增强_build_table_from_cells方法，使其能处理字典类型的表格对象
   - 增强_detect_merged_cells方法，使其能处理不同类型的表格对象

3. 表格区域标记修复:
   - 增强_mark_table_regions方法，使其能处理不同格式的表格对象
   - 正确计算表格边界，无论表格对象类型如何

注意事项
-------

1. 这些修复适用于各种版本的PyMuPDF，但建议使用最新版本以获取最佳效果
2. 某些复杂的PDF文档中的表格可能仍然无法完美提取，这取决于PDF的结构
3. 如果遇到其他问题，请查看错误信息，可能需要进一步修改代码
"""
