[![Tests](https://github.com/sir-go/reports-xls/actions/workflows/python-app.yml/badge.svg)](https://github.com/sir-go/reports-xls/actions/workflows/python-app.yml)

## TeleTime billing DB to MS Excel report generator

The script requests data from the MySQL DB and generates multi-page 
XLS report with formulas and fields freezing.

### Install

To virtualenv

```bash
virtualenv venv
source ./venv/bin/activate
pip install -r requirements.txt
```
or build a Docker image

| build-arg | meaning           | default  |
|-----------|-------------------|----------|
| UID       | running user ID   | 1000     |
| GID       | running group ID  | 1000     |
| OUT       | report saving dir | /app/out |

```bash
docker build --build-arg UID=$(id -u) --build-arg GID=$(id -g) . -t reports-xls
```
___
### Config

Env variables

| variable     | meaning           |
|--------------|-------------------|
| REP_HOST     | DB host           |
| REP_USERNAME | DB username       |
| REP_PASSWORD | DB password       |
| REP_DB       | DB name           |
| REP_OUT      | report saving dir |

___
### Test

```bash
python -m pytest
```

___
### Run

#### Standalone

```bash
python make_report.py
```
#### or Docker container (variables are in the `.env` file here)

```bash
docker run --rm -it --env-file .env -v ${PWD}/out:/app/out reports-xls
```
