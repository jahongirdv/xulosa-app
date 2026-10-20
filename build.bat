@echo off
chcp 65001 >nul
title Сборка Xulosa Generator

echo ================================================
echo  1. Активация виртуального окружения...
echo ================================================
call .venv\Scripts\activate.bat

echo.
echo ================================================
echo  2. Установка зависимостей...
echo ================================================
pip install -r requirements.txt

echo.
echo ================================================
echo  3. Очистка старых файлов сборки...
echo ================================================
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "XulosaGenerator.spec" del /q "XulosaGenerator.spec"

echo.
echo ================================================
echo  4. Запуск PyInstaller...
echo ================================================
pyinstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name "XulosaGenerator" ^
  --icon "images.ico" ^
  --add-data "shablon.docx;." ^
  --add-data "images.ico;." ^
  main.py

echo.
echo ================================================
echo  ГОТОВО! Ваш файл здесь: dist\XulosaGenerator.exe
echo ================================================
pause