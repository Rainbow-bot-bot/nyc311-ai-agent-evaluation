@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -B -X utf8 "评测数据处理.py" --sync-scores
if errorlevel 1 exit /b 1
python -B -X utf8 "评测数据处理.py" --check-scores
if errorlevel 1 exit /b 1
echo 已同步并核对现行S评分与查证数据。报告表图按资料/复现说明.md维护。
exit /b 0
