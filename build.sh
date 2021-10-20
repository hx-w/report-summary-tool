pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
rm -rf build/ dist/
rm -f main.spec
pyinstaller -D -w -i static/icon.ico main.py
rm -rf build/
rm -f main.spec

if [ -f "dist/main.exe" ]; then
    cp config.yml dist/main.exe env/06_工作周报/
fi