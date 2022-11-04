[![Tests](https://github.com/sir-go/reports-xls/actions/workflows/python-app.yml/badge.svg)](https://github.com/sir-go/reports-xls/actions/workflows/python-app.yml)

## TeleTime billing DB to MS Excel report generator

The script requests data from the MySQL DB and generates multi-page XLS report with formulas and fields freezing.

### Install
```bash
virtualenv venv
source ./venv/bin/activate
pip install -r requirements.txt
```

### Config
`config.py` must contain `db_conf` dict:
```python
db_conf = dict(
    host='',    # DB host
    user='',    # DB username
    passwd='',  # DB password
    db=''       # DB name
)
```

### Run
```bash
python make_report.py
```
