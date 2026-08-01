@echo off
chcp 65001 >nul
echo ============================================
echo  Outlook PST 搜索工具 - 打包脚本
echo  需要在 Windows 电脑上，安装了 Python 3.10+ 后运行
echo ============================================

echo [1/3] 安装依赖库 ...
pip install -r requirements.txt
if errorlevel 1 (
    echo 依赖安装失败，请检查网络或 Python 环境
    pause
    exit /b 1
)

echo [2/3] 使用 PyInstaller 打包为单文件 exe ...
pyinstaller --onefile --windowed --name "PST邮件搜索工具" --add-data "assets;assets" pst_search_tool.py

echo [3/3] 打包完成！
echo 生成的 exe 文件在 dist\PST邮件搜索工具.exe
echo 将该 exe 文件直接拷贝给同事即可使用，无需安装 Python。
pause
