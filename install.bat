@echo off
chcp 65001 >nul
echo ====================================
echo   教案合规性检测智能体 - 依赖安装
echo ====================================
echo.

echo [1/3] 检查Python版本...
python --version
if errorlevel 1 (
    echo ❌ 错误: 未找到Python，请先安装Python 3.8或更高版本
    pause
    exit /b 1
)
echo.

echo [2/3] 升级pip...
python -m pip install --upgrade pip
echo.

echo [3/3] 安装依赖包...
pip install -r requirements.txt
echo.

if errorlevel 0 (
    echo ====================================
    echo   ✅ 所有依赖安装完成！
    echo ====================================
    echo.
    echo 现在可以运行程序了：
    echo   python agent.py --file 教案测试2/string.md
    echo.
) else (
    echo ====================================
    echo   ❌ 安装过程中出现错误
    echo ====================================
    echo.
    echo 请尝试手动安装：
    echo   pip install langchain langchain-core langchain-openai openai python-dotenv Pillow markdown beautifulsoup4 pydantic
    echo.
)

pause
