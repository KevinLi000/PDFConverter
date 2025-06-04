"""
复杂表格格式增强模块 - 处理复杂表格格式和嵌套表格
"""

import os
import traceback
import types
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Pt, Cm

def enhance_complex_table_handling(converter):
    """
    增强复杂表格处理能力
    
    参数:
        converter: PDF转换器实例
        
    返回:
        布尔值，表示是否成功应用增强功能
    """
    print("正在应用复杂表格格式增强...")
    
    try:
        # 为转换器添加复杂表格处理功能
        
        # 增强表格单元格检测方法
        if hasattr(converter, '_validate_and_fix_table_data'):
            original_validate = converter._validate_and_fix_table_data
              def enhanced_validate_and_fix_table_data(self, table_data, merged_cells=None):
                """增强的表格数据验证方法，更好地处理复杂单元格内容"""
                # 先使用原始方法进行基本验证
                fixed_data, fixed_merged = original_validate(table_data, merged_cells)
                
                # 增强处理 - 确保所有单元格内容都被正确处理
                if fixed_data:
                    for i in range(len(fixed_data)):
                        for j in range(len(fixed_data[i])):
                            cell_content = fixed_data[i][j]
                            
                            # 处理嵌套结构
                            if isinstance(cell_content, dict):
                                # 如果单元格内容是字典，提取文本内容
                                if 'text' in cell_content:
                                    fixed_data[i][j] = cell_content['text']
                                elif 'spans' in cell_content:
                                    # 合并所有spans中的文本
                                    text = ""
                                    for span in cell_content['spans']:
                                        if 'text' in span:
                                            text += span['text']
                                    fixed_data[i][j] = text
                            
                            # 确保换行符保留
                            if isinstance(fixed_data[i][j], str):
                                # 统一换行符格式
                                fixed_data[i][j] = fixed_data[i][j].replace('\\n', '\n')
                
                return fixed_data, fixed_merged
            
            # 替换原始方法
            converter._validate_and_fix_table_data = types.MethodType(enhanced_validate_and_fix_table_data, converter)
        
        # 为转换器添加更精确的表格样式应用方法
        def apply_advanced_table_style(self, table, style_info=None):
            """应用高级表格样式，确保精确保留表格格式"""
            try:
                from docx.oxml import OxmlElement, parse_xml
                from docx.oxml.ns import nsdecls, qn
                
                # 设置表格基本样式
                table.style = 'Table Grid'
                
                # 设置表格对齐方式
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                # 设置表格边框 - 使用更明确的边框设置
                tbl_pr = table._tbl.xpath('./w:tblPr')[0]
                
                # 定义边框
                border_size = 8
                border_color = "000000"  # 黑色
                
                # 创建边框XML
                borders_xml = f'''
                <w:tblBorders {nsdecls("w")}>
                  <w:top w:val="single" w:sz="{border_size}" w:space="0" w:color="{border_color}"/>
                  <w:left w:val="single" w:sz="{border_size}" w:space="0" w:color="{border_color}"/>
                  <w:bottom w:val="single" w:sz="{border_size}" w:space="0" w:color="{border_color}"/>
                  <w:right w:val="single" w:sz="{border_size}" w:space="0" w:color="{border_color}"/>
                  <w:insideH w:val="single" w:sz="{border_size}" w:space="0" w:color="{border_color}"/>
                  <w:insideV w:val="single" w:sz="{border_size}" w:space="0" w:color="{border_color}"/>
                </w:tblBorders>
                '''
                
                # 删除已存在的边框设置
                existing_borders = tbl_pr.xpath('./w:tblBorders')
                for border in existing_borders:
                    tbl_pr.remove(border)
                
                # 添加新的边框设置
                tbl_pr.append(parse_xml(borders_xml))
                
                # 设置表格布局 - 使用固定宽度而非自动调整
                tbl_layout = OxmlElement('w:tblLayout')
                tbl_layout.set(qn('w:type'), 'fixed')
                
                # 删除现有布局设置
                existing_layout = tbl_pr.xpath('./w:tblLayout')
                for layout in existing_layout:
                    tbl_pr.remove(layout)
                
                # 添加新的布局设置
                tbl_pr.append(tbl_layout)
                
                # 禁用自动调整
                table.autofit = False
                
                # 设置每个单元格的格式
                for row in table.rows:
                    for cell in row.cells:
                        # 设置垂直对齐
                        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                        
                        # 设置单元格边框
                        tc_pr = cell._element.get_or_add_tcPr()
                        
                        # 创建单元格边框XML
                        cell_borders_xml = f'''
                        <w:tcBorders {nsdecls("w")}>
                          <w:top w:val="single" w:sz="{border_size}" w:space="0" w:color="{border_color}"/>
                          <w:left w:val="single" w:sz="{border_size}" w:space="0" w:color="{border_color}"/>
                          <w:bottom w:val="single" w:sz="{border_size}" w:space="0" w:color="{border_color}"/>
                          <w:right w:val="single" w:sz="{border_size}" w:space="0" w:color="{border_color}"/>
                        </w:tcBorders>
                        '''
                        
                        # 删除现有边框
                        existing_cell_borders = tc_pr.xpath('./w:tcBorders')
                        for border in existing_cell_borders:
                            tc_pr.remove(border)
                        
                        # 添加新的边框
                        tc_pr.append(parse_xml(cell_borders_xml))
                        
                        # 设置单元格内边距
                        margins_xml = f'''
                        <w:tcMar {nsdecls("w")}>
                          <w:top w:w="100" w:type="dxa"/>
                          <w:left w:w="100" w:type="dxa"/>
                          <w:bottom w:w="100" w:type="dxa"/>
                          <w:right w:w="100" w:type="dxa"/>
                        </w:tcMar>
                        '''
                        
                        # 删除现有内边距
                        existing_margins = tc_pr.xpath('./w:tcMar')
                        for margin in existing_margins:
                            tc_pr.remove(margin)
                        
                        # 添加新的内边距
                        tc_pr.append(parse_xml(margins_xml))
                        
                        # 优化段落间距
                        for paragraph in cell.paragraphs:
                            if paragraph.text.strip():
                                paragraph.space_before = Pt(0)
                                paragraph.space_after = Pt(0)
                                
                                # 确保段落中的文本格式一致
                                if paragraph.runs:
                                    for run in paragraph.runs:
                                        # 设置基本字体
                                        run.font.name = "Arial"
                                        run.font.size = Pt(10)
                
                return True
            except Exception as e:
                print(f"应用高级表格样式时出错: {e}")
                traceback.print_exc()
                return False
        
        # 添加方法到转换器
        converter.apply_advanced_table_style = types.MethodType(apply_advanced_table_style, converter)
        
        # 增强表格检测能力
        if hasattr(converter, '_detect_table_style'):
            original_detect_style = converter._detect_table_style
            
            def enhanced_detect_table_style(self, block, page):
                """增强的表格样式检测方法"""
                try:
                    # 尝试使用enhanced_table_style模块
                    from enhanced_table_style import detect_table_style
                    return detect_table_style(block, page)
                except ImportError:
                    # 回退到原始方法
                    return original_detect_style(self, block, page)
                except Exception as e:
                    print(f"增强表格样式检测时出错: {e}")
                    # 回退到原始方法
                    return original_detect_style(self, block, page)
            
            # 替换原始方法
            converter._detect_table_style = types.MethodType(enhanced_detect_table_style, converter)
        
        print("复杂表格格式增强功能已成功应用")
        return True
    
    except Exception as e:
        print(f"应用复杂表格格式增强时出错: {e}")
        traceback.print_exc()
        return False
