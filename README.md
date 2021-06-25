# Auto-fill_horarios
to fill UPM required horarios register automatically

# Dependecies
```
pip install pandas
pip install openpyxl
pip install holidays-es
pip install pillow
pip install BeautifulSoup4
```
# INPUT information
change your input information in `./auto-fill_horarios.py`
```
# fill basic info
NAME='ZIQI DENG'
NIF='EDXXXXXX'
YEAR=2020
HORA_ENTRADA_MORNING = "09:00:00"
HORA_SALIDA_MORNING = "13:00:00"
HORA_ENTRADA_AFTERNOON = "14:00:00"
HORA_SALIDA_AFTERNOON = "17:30:00"
TOTAL_HORAS = 7.5

ANNUAL_LEAVES = [
    # start-date, end date Year-Month-Day
    ["2020-01-04", "2020-01-10"], 
    ["2020-02-04", "2020-02-14"],
]
TEMPLATE = './registro_jornada_laboral_template.xlsx'
```
# Usage
```
python auto-fill_horarios.py
```

# Required file
model format is according to `registro_jornada_laboral_template.xlsx`. Check output file example with `registro_jornada_laboral_ENERO2020`

# Style
open output file with libre office may cause partially loss of style 