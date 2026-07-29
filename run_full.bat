@echo off
cd /d "%~dp0"
py -m pip install -r requirements.txt
py -m unittest discover -s tests -v
py scraper.py --template template.xlsx --config config.json --limit 0
pause
