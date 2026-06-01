@echo off
title AVE Seguimiento PRO
cd /d %~dp0
python -m pip install -r requirements.txt
streamlit run app.py
pause
