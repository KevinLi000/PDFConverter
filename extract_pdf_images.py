"""
图像提取测试工具 - 检查PDF中的图像并提取到指定目录
"""

import os
import sys
import fitz  # PyMuPDF
import time
from PIL import Image

def extract_images_from_pdf(pdf_path, output_dir=None, min_size=100):
    """
    从PDF中提取所有图像
    
    参数:
        pdf_path: PDF文件路径
        output_dir: 输出目录，默认为PDF所在目录下的images文件夹
        min_size: 最小图像尺寸(像素)，小于此尺寸的图像将被忽略
        
    返回:
        提取的图像数量
    """
    if not output_dir:
        pdf_dir = os.path.dirname(os.path.abspath(pdf_path))
        pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_dir = os.path.join(pdf_dir, f"{pdf_name}_images")
    
    os.makedirs(output_dir, exist_ok=True)
    print(f"图像将保存到: {output_dir}")
    
    # 打开PDF
    doc = fitz.open(pdf_path)
    total_images = 0
    extracted_xrefs = set()
    
    # 遍历所有页面
    for page_idx in range(len(doc)):
        page = doc[page_idx]
        print(f"处理页面 {page_idx+1}/{len(doc)}")
        
        # 方法1: 使用get_images()
        try:
            image_list = page.get_images()
            for img_idx, img in enumerate(image_list):
                xref = img[0]
                
                # 跳过已提取的图像
                if xref in extracted_xrefs:
                    continue
                
                try:
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    image_ext = base_image["ext"]
                    
                    # 保存图像
                    img_filename = f"page{page_idx+1}_xref{xref}.{image_ext}"
                    img_path = os.path.join(output_dir, img_filename)
                    
                    with open(img_path, "wb") as img_file:
                        img_file.write(image_bytes)
                    
                    # 检查图像尺寸
                    try:
                        with Image.open(img_path) as img:
                            width, height = img.size
                            if width < min_size or height < min_size:
                                os.remove(img_path)
                                print(f"  跳过小图像: {img_filename} ({width}x{height})")
                                continue
                    except Exception:
                        pass
                    
                    print(f"  提取图像: {img_filename}")
                    total_images += 1
                    extracted_xrefs.add(xref)
                except Exception as e:
                    print(f"  无法提取图像 xref={xref}: {e}")
        except Exception as e:
            print(f"  使用get_images()方法出错: {e}")
        
        # 方法2: 通过页面块提取
        try:
            blocks = page.get_text("dict")["blocks"]
            for block_idx, block in enumerate(blocks):
                if block["type"] == 1:  # 图像块
                    try:
                        # 获取图像xref
                        xref = block.get("xref", 0)
                        
                        # 如果没有xref或已提取，尝试通过裁剪获取
                        if xref == 0 or xref in extracted_xrefs:
                            bbox = block["bbox"]
                            clip_rect = fitz.Rect(bbox)
                            
                            # 使用高分辨率
                            zoom = 4.0
                            matrix = fitz.Matrix(zoom, zoom)
                            pix = page.get_pixmap(matrix=matrix, clip=clip_rect, alpha=False)
                            
                            # 保存图像
                            img_filename = f"page{page_idx+1}_block{block_idx}.png"
                            img_path = os.path.join(output_dir, img_filename)
                            pix.save(img_path)
                            
                            # 检查图像尺寸
                            try:
                                with Image.open(img_path) as img:
                                    width, height = img.size
                                    if width < min_size or height < min_size:
                                        os.remove(img_path)
                                        print(f"  跳过小图像: {img_filename} ({width}x{height})")
                                        continue
                            except Exception:
                                pass
                            
                            print(f"  提取图像块: {img_filename}")
                            total_images += 1
                    except Exception as e:
                        print(f"  无法提取图像块: {e}")
        except Exception as e:
            print(f"  通过页面块提取出错: {e}")
    
    print(f"\n共提取 {total_images} 个图像到 {output_dir}")
    return total_images

def main():
    if len(sys.argv) < 2:
        print("使用方法: python extract_pdf_images.py <pdf文件路径> [输出目录] [最小图像尺寸]")
        return 1
    
    pdf_path = sys.argv[1]
    if not os.path.exists(pdf_path):
        print(f"错误: 找不到文件 {pdf_path}")
        return 1
    
    output_dir = None
    if len(sys.argv) >= 3:
        output_dir = sys.argv[2]
    
    min_size = 100
    if len(sys.argv) >= 4:
        try:
            min_size = int(sys.argv[3])
        except ValueError:
            print(f"警告: 无效的最小尺寸值 '{sys.argv[3]}'，使用默认值 100")
    
    start_time = time.time()
    try:
        extract_images_from_pdf(pdf_path, output_dir, min_size)
        end_time = time.time()
        print(f"处理耗时: {end_time - start_time:.2f} 秒")
        return 0
    except Exception as e:
        print(f"图像提取过程中出错: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
