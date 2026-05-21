@echo off
echo Starting Nexus Governance Intelligence...
call venv\Scripts\activate
venv\Scripts\python.exe -m streamlit run app/main.py
