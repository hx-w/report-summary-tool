pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
pyinstaller -F -i static/icon.ico main.py
rmdir build/
