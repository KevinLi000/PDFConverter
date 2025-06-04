# 换行符处理增强 (Line Break Enhancement)

此模块用于改进PDF转Word过程中换行符的识别和处理，确保PDF文档中的换行被正确保留在Word文档中。

## 功能特点

- 精确识别和保留PDF中的换行符
- 智能区分段落换行和行内换行
- 基于行间距、缩进和文本内容自动判断段落结构
- 处理各种特殊格式的换行标记
- 无缝集成到现有PDF转换流程中

## 安装与使用

1. 确保已安装必要的依赖项：

```bash
pip install PyMuPDF python-docx
```

2. 将`line_break_enhancement.py`放置到您的项目目录中

3. 在代码中引用此模块：

```python
from line_break_enhancement import enhance_line_break_handling, apply_line_break_enhancement

# 方法1：直接应用到PDF转换器
converter = EnhancedPDFConverter()
enhance_line_break_handling(converter)

# 方法2：使用便捷函数
converter = EnhancedPDFConverter()
apply_line_break_enhancement(converter)

# 方法3：通过all_pdf_fixes_integrator集成
from all_pdf_fixes_integrator import integrate_all_fixes
converter = EnhancedPDFConverter()
converter = integrate_all_fixes(converter)  # 已包含换行符处理增强
```

## 测试和验证

可以使用提供的测试脚本验证换行符处理效果：

```bash
python test_line_break.py path/to/your/pdf/file.pdf
```

此命令将创建一个测试文档，展示换行符处理的效果。

## 换行符处理策略

此模块使用以下策略识别和处理换行符：

1. **显式换行符检测**：识别文本中的`\n`、`\r\n`等显式换行符
2. **行间距分析**：通过分析行间距判断是否为段落换行
3. **缩进检测**：通过分析行的缩进判断段落结构
4. **文本内容分析**：通过分析句子结构（如句号后跟大写字母）判断段落边界
5. **样式检测**：根据字体大小、粗细等判断标题和段落边界

## 集成到现有项目

若要将此模块集成到现有项目中，您可以修改`all_pdf_fixes_integrator.py`，确保其中包含了换行符处理增强的调用。最新版本的集成器已经包含了此功能。

## 主要修复的问题

1. PDF转换到Word时换行符丢失
2. 段落结构错误（多个段落被合并或单个段落被拆分）
3. 列表项换行处理不当
4. 特殊格式文本（如代码块）的换行符丢失

## 注意事项

- 此模块与其他格式保留增强功能（如图像恢复）完全兼容
- 使用"maximum"格式保留级别可获得最佳的换行符处理效果
- 对于包含复杂排版的PDF文件，可能需要微调参数以获得最佳效果
