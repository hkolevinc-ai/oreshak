@echo off
cd /d "%~dp0"
py -m pip install -r requirements.txt
py scraper.py --template template.xlsx --config config.json --limit 0
pause
