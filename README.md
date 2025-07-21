# python-wa
$env:PLAYWRIGHT_BROWSERS_PATH = ".\browser"
pip freeze > requirements.txt
pip install -r requirements.txt
playwright install chronium