mode con cols=60 lines=20
color 2f
@echo off
@cls
title "国家税务总局发票查验"
call .\venv\Scripts\activate.bat

@echo.
echo 查验发票注意事项
@echo.
echo 1, 查验前，关闭所有的浏览器
@echo.
echo 2, 查验中，请勿操作其他软件
echo    "验证码需要手工输入"
echo    "验证码输入完毕后点击<查验>按钮"
@echo.
echo 3, 结束查验前，请关闭程序打开的浏览器
@echo.
echo 4, 结束查验，请关闭本程序
@echo.
echo 5, 按任意键开始发票查验
@echo.

pause
python code07.py

call .\venv\Scripts\deactivate.bat

@echo on