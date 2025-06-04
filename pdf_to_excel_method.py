def pdf_to_excel(self, method="advanced"):
    """
    将PDF转换为Excel文件，保留格式
    
    参数:
        method (str): 转换方法，可选值为 "basic", "standard", "advanced"
    
    返回:
        str: 输出文件路径
    """
    try:
        import os
        import pandas as pd
        import tempfile
        
        # 确保输出目录存在
        os.makedirs(self.output_dir, exist_ok=True)
        
        # 创建输出文件路径
        input_filename = os.path.basename(self.pdf_path)
        base_name = os.path.splitext(input_filename)[0]
        output_path = os.path.join(self.output_dir, f"{base_name}.xlsx")
        
        # 使用PyMuPDF打开PDF
        pdf_document = fitz.open(self.pdf_path)
        
        # 创建一个Excel写入器
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 处理每一页
            for page_num in range(len(pdf_document)):
                # 提取表格 - 使用tabula或其他可用方法
                try:
                    # 首先尝试使用tabula
                    import tabula
                    tables = tabula.read_pdf(
                        self.pdf_path, 
                        pages=page_num + 1,  # tabula使用1-based页码
                        multiple_tables=True,
                        guess=True,
                        stream=method != "advanced",
                        lattice=method == "advanced"
                    )
                    
                    if not tables:
                        # 如果tabula没有检测到表格，尝试使用PyMuPDF的表格检测
                        if hasattr(self, '_extract_tables'):
                            tables_from_pymupdf = self._extract_tables(pdf_document, page_num)
                            # 将PyMuPDF格式转换为pandas DataFrame
                            if tables_from_pymupdf:
                                # 处理自定义表格格式...
                                pass
                except ImportError:
                    # 如果tabula不可用，使用PyMuPDF提取文本并尝试解析表格
                    if hasattr(self, '_extract_tables'):
                        tables_from_pymupdf = self._extract_tables(pdf_document, page_num)
                        # 处理自定义表格格式...
                    else:
                        # 回退到基本文本提取方法
                        page = pdf_document[page_num]
                        text = page.get_text("text")
                        # 尝试使用文本构建基本表格
                        tables = []
                        # ... 基本表格解析逻辑 ...
                except Exception as e:
                    print(f"提取表格错误 (页面 {page_num+1}): {e}")
                    tables = []
                
                # 将表格写入Excel工作表
                if tables:
                    for i, table in enumerate(tables):
                        if isinstance(table, pd.DataFrame):
                            sheet_name = f"Page{page_num+1}_Table{i+1}"
                            if len(sheet_name) > 31:  # Excel工作表名称长度限制
                                sheet_name = sheet_name[:31]
                            table.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    # 如果没有检测到表格，创建一个空工作表
                    sheet_name = f"Page{page_num+1}"
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    pd.DataFrame().to_excel(writer, sheet_name=sheet_name)
                    
                    # 提取页面文本并添加到工作表
                    page = pdf_document[page_num]
                    text = page.get_text("text")
                    
                    # 获取工作表
                    worksheet = writer.sheets[sheet_name]
                    
                    # 按行分割文本并写入
                    lines = text.split('\n')
                    for row_idx, line in enumerate(lines):
                        worksheet.cell(row=row_idx+1, column=1, value=line)
                    
                    # 调整列宽
                    worksheet.column_dimensions['A'].width = 100
        
        # 关闭PDF
        pdf_document.close()
        
        return output_path
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"PDF到Excel转换失败: {e}")
        
        # 创建一个基本的Excel文件作为后备方案
        try:
            import pandas as pd
            
            # 创建输出文件路径
            input_filename = os.path.basename(self.pdf_path)
            base_name = os.path.splitext(input_filename)[0]
            output_path = os.path.join(self.output_dir, f"{base_name}.xlsx")
            
            # 创建一个包含错误信息的DataFrame
            df = pd.DataFrame({
                "错误": [f"PDF到Excel转换失败: {e}"],
                "提示": ["请尝试使用不同的转换方法或联系开发人员"]
            })
            
            # 保存到Excel
            df.to_excel(output_path, index=False)
            
            return output_path
            
        except Exception as backup_err:
            print(f"创建备用Excel文件也失败: {backup_err}")
            raise e
