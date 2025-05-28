@echo off
echo 正在安装所需依赖...
pip install -r requirements.txt
echo.
echo 启动PDF转Word转换器...
python pdf_to_word_cn.py
pause
