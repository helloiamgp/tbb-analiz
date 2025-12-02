@echo off
title TBB Yanit Sistemi - Kurulum

echo.
echo ========================================================
echo     TBB Yanit Sistemi - EXE Olusturucu
echo     Aytemiz Yatirim Bankasi A.S.
echo ========================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [HATA] Python bulunamadi!
    echo Lutfen Python 3.8+ yukleyin: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/4] Kutuphaneler yukleniyor...
pip install python-docx pandas openpyxl PyMuPDF pyinstaller --quiet --disable-pip-version-check

echo.
echo [2/4] EXE olusturuluyor (2-3 dakika surebilir)...

python -m PyInstaller --onefile --windowed --name=TBB_Yanit_Sistemi --clean --noconfirm --add-data "logo.png;." main.py

echo.
echo [3/4] Dagitim klasoru hazirlaniyor...

if not exist "Dagitim" mkdir Dagitim
copy dist\TBB_Yanit_Sistemi.exe Dagitim\ >nul
if exist musteri_listesi.xlsx copy musteri_listesi.xlsx Dagitim\ >nul

echo TBB Yazi Otomatik Yanitlama Sistemi > Dagitim\KULLANIM.txt
echo =================================== >> Dagitim\KULLANIM.txt
echo. >> Dagitim\KULLANIM.txt
echo KULLANIM: >> Dagitim\KULLANIM.txt
echo 1. TBB_Yanit_Sistemi.exe dosyasini calistirin >> Dagitim\KULLANIM.txt
echo 2. Musteri Listesi Yukle ile Excel dosyanizi secin >> Dagitim\KULLANIM.txt
echo 3. Tekli veya Toplu islem yapin >> Dagitim\KULLANIM.txt
echo. >> Dagitim\KULLANIM.txt
echo Aytemiz Yatirim Bankasi A.S. - 2025 >> Dagitim\KULLANIM.txt

echo.
echo [4/4] Temizlik...
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q *.spec 2>nul

echo.
echo ========================================================
echo   TAMAMLANDI!
echo   Dagitim klasorunu kullanicilara paylasabilirsiniz.
echo ========================================================
echo.

explorer Dagitim
pause
