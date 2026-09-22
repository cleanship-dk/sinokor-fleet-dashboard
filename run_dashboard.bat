@echo off
REM Launch the CLS fleet targeting dashboard in your browser.
cd /d "%~dp0"
echo Checking dependencies...
python -m pip install --quiet --disable-pip-version-check -r requirements.txt
echo Starting dashboard...
python -m streamlit run app.py
pause
