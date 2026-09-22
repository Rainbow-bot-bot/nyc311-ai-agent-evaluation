@echo off
chcp 65001 >nul
cd /d "%~dp0"

python -B -X utf8 "评测数据处理.py"
if errorlevel 1 goto :fail

python -B -X utf8 "生成分析表.py"
if errorlevel 1 goto :fail

echo.
echo 已完成：项目2_AI评测分析.xlsx 已重建并验收。
goto :end

:fail
echo.
echo 重建失败。正式Excel未通过验收时不会被替换，请查看上面的具体错误。

:end
pause
