mode con cols=60 lines=20
color 2f
@echo off
@cls
title "����˰���ַܾ�Ʊ����"
call .\venv\Scripts\activate.bat

@echo.
echo ���鷢Ʊע������
@echo.
echo 1, ����ǰ���ر����е������
@echo.
echo 2, �����У���������������
echo    "��֤����Ҫ�ֹ�����"
echo    "��֤��������Ϻ���<����>��ť"
@echo.
echo 3, ��������ǰ����رճ���򿪵������
@echo.
echo 4, �������飬��رձ�����
@echo.
echo 5, ���������ʼ��Ʊ����
@echo.

pause
python code07.py

call .\venv\Scripts\deactivate.bat

@echo on