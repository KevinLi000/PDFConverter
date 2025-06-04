# 图像恢复增强模块 (Image Recovery Enhancement)

此模块用于解决PDF转Word过程中图像丢失的问题，通过多种方法提取和恢复PDF文档中的图像。

## 功能特点

- 使用多种图像提取方法，确保最大程度保留PDF中的图像
- 智能选择最佳提取结果，优化图像质量
- 自动处理损坏或低质量图像
- 无缝集成到现有PDF转换流程中
- 支持各种复杂PDF文档格式

## 安装与使用

1. 确保已安装必要的依赖项：

```bash
pip install PyMuPDF python-docx Pillow
```

2. 将`image_recovery_enhancement.py`放置到您的项目目录中

3. 在代码中引用此模块：

```python
from image_recovery_enhancement import enhance_image_extraction, apply_image_recovery

# 方法1：集成到PDF转换器
converter = EnhancedPDFConverter()
enhance_image_extraction(converter)

# 方法2：使用all_pdf_fixes_integrator
from all_pdf_fixes_integrator import integrate_all_fixes
converter = EnhancedPDFConverter()
converter = integrate_all_fixes(converter)  # 已包含图像恢复增强
```

## 测试和验证

可以使用提供的测试脚本验证图像恢复功能：

```bash
python test_pdf_image_recovery.py path/to/your/pdf/file.pdf
```

此命令将转换PDF并显示图像保留统计信息。

## 图像恢复策略

此模块使用以下策略提取和恢复图像：

1. **xref提取**: 通过图像引用直接提取嵌入图像
2. **get_image方法**: 使用PyMuPDF的get_images和extract_image方法
3. **高分辨率裁剪**: 使用高分辨率对页面区域进行裁剪
4. **扩展边界框**: 扩大检测区域以捕获可能被错误裁剪的图像
5. **整页裁剪**: 对复杂页面使用全页面提取再裁剪的方法

## 集成到现有项目

若要将此模块集成到现有项目中，请运行：

```bash
python integrate_image_recovery.py
```

这将自动将图像恢复增强集成到PDF转换流程中。

## 主要修复的问题

1. 部分或全部图像在转换过程中丢失
2. 图像质量降低或变形
3. 透明图像处理不当
4. 嵌入图像未被识别

## 注意事项

- 此模块可能会生成临时图像文件，它们将保存在转换器的`temp_dir`目录下
- 使用"maximum"格式保留级别将获得最佳的图像恢复效果
- 对于大型或复杂的PDF文件，处理时间可能会增加
