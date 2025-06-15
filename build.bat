@echo off
setlocal enabledelayedexpansion

:: Конфигурация сборки
set "PROJECT_NAME=DIPLOM_PROG"
set "UI_FILE=untitled_2.ui"
set "ICON_FILE=app_icon.ico"
set "MAIN_SCRIPT=%PROJECT_NAME%.py"
set "TARGET_DIR=C:\Users\Dimaf\Desktop\Diplom"
set "REQUIREMENTS=requirements.txt"

:: Проверка наличия Python
where python >nul 2>nul
if errorlevel 1 (
    echo ОШИБКА: Python не найден в системе!
    echo Убедитесь, что Python установлен и добавлен в PATH.
    pause
    exit /b 1
)

:: Проверка наличия PyInstaller
echo Проверка наличия PyInstaller...
python -c "import pyinstaller" >nul 2>nul
if errorlevel 1 (
    echo PyInstaller не найден. Установка...
    pip install pyinstaller
    if errorlevel 1 (
        echo ОШИБКА: Не удалось установить PyInstaller!
        echo Попробуйте установить вручную: pip install pyinstaller
        pause
        exit /b 1
    )
)

:: Установка зависимостей проекта
echo Установка зависимостей проекта...
if exist "%REQUIREMENTS%" (
    pip install -r "%REQUIREMENTS%"
) else (
    echo Файл требований "%REQUIREMENTS%" не найден. Установка основных зависимостей...
    pip install pandas PyQt5 numpy openpyxl xlrd python-docx scipy scikit-learn
)

:: Проверка целевой директории
if not exist "%TARGET_DIR%" (
    echo Создание целевой директории: %TARGET_DIR%
    mkdir "%TARGET_DIR%"
    if errorlevel 1 (
        echo ОШИБКА: Не удалось создать целевую директорию!
        pause
        exit /b 1
    )
)

:: Проверка необходимых файлов
if not exist "%MAIN_SCRIPT%" (
    echo ОШИБКА: Основной скрипт "%MAIN_SCRIPT%" не найден!
    pause
    exit /b 1
)

if not exist "%UI_FILE%" (
    echo ОШИБКА: Файл интерфейса "%UI_FILE%" не найден!
    pause
    exit /b 1
)

if not exist "%ICON_FILE%" (
    echo ВНИМАНИЕ: Файл иконки "%ICON_FILE%" отсутствует. Будет использована иконка по умолчанию.
    set "ICON_OPTION="
) else (
    set "ICON_OPTION=--icon=%ICON_FILE%"
)

:: Очистка предыдущих артефактов
if exist "build" rmdir /s /q "build"
del /q "%PROJECT_NAME%.spec" 2>nul

:: Удаление старого файла в целевой директории
if exist "%TARGET_DIR%\%PROJECT_NAME%.exe" (
    echo Удаление предыдущей версии...
    del /q "%TARGET_DIR%\%PROJECT_NAME%.exe"
)

:: Команда сборки с явным указанием всех зависимостей
echo Запуск сборки с включением зависимостей...
python -m PyInstaller ^
    --onefile ^
    --name "%PROJECT_NAME%" ^
    --add-data "%UI_FILE%;." ^
    --windowed ^
    %ICON_OPTION% ^
    --clean ^
    --distpath "%TARGET_DIR%" ^
    --workpath "build" ^
    --noconfirm ^
    --hidden-import pandas ^
    --hidden-import PyQt5 ^
    --hidden-import numpy ^
    --hidden-import openpyxl ^
    --hidden-import xlrd ^
    --hidden-import docx ^
    --hidden-import docx.shared ^
    --hidden-import docx.enum ^
    --hidden-import lxml ^
    --hidden-import scipy ^
    --hidden-import scipy.signal ^
    --hidden-import scipy.fftpack ^
    --hidden-import scipy.interpolate ^
    --hidden-import scipy.special ^
    --hidden-import scipy.stats ^
    --hidden-import sklearn ^
    --hidden-import sklearn.ensemble ^
    --hidden-import sklearn.tree ^
    --hidden-import sklearn.neighbors ^
    --hidden-import sklearn.svm ^
    --hidden-import sklearn.linear_model ^
    --hidden-import sklearn.cluster ^
    --hidden-import sklearn.model_selection ^
    --hidden-import sklearn.preprocessing ^
    --hidden-import sklearn.metrics ^
    "%MAIN_SCRIPT%"

:: Проверка результата сборки
if errorlevel 1 (
    echo.
    echo ОШИБКА сборки! Проверьте сообщения выше.
    echo Возможные причины:
    echo 1. Недостающие зависимости Python
    echo 2. Ошибки в коде приложения
    pause
    exit /b 1
)

if exist "%TARGET_DIR%\%PROJECT_NAME%.exe" (
    echo.
    echo Сборка успешно завершена!
    echo Исполняемый файл сохранен в: %TARGET_DIR%\%PROJECT_NAME%.exe
    echo Размер файла: %~z0
) else (
    echo.
    echo ОШИБКА: Исполняемый файл не создан!
)

echo.
echo Нажмите любую клавишу для выхода...
pause >nul