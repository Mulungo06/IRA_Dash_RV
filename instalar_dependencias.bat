@echo off
echo Instalando dependencias necessarias...
echo.

pip install python-docx
pip install streamlit
pip install pandas
pip install openpyxl
pip install numpy

echo.
echo ========================================
echo Instalacao concluida!
echo ========================================
echo.
echo Agora execute: streamlit run app_iras_completo.py
echo.
pause
