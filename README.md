# FB-Leads-ToDb

Make insert statements from Facebook Leads form stored in xlsx file.

## Requirements

- [Python 3](https://www.python.org/downloads/)
- An xlsx file with data

## Setup

```
pip install openpyxl
pip install termcolor
```

## Usage

```
python .\leadsToDb.py -c "file\path.xlsx" -m "Manual import date dd/mm"
```
For help:

```
python .\leadsToDb.py -h
```