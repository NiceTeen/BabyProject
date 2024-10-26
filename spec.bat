@echo off
call G:\AllVenv\BabyProject\Scripts\activate
pyinstaller -F -w -i ico.ico main.py
pause