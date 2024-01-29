# Spreadsheet Parser

Use python / pandas and argparse to parse an XLS spreadsheet for two columns of a given sheet.

Quick, simple, dirty swiss army knife. Outputs a 2 column csv.

## Install / Setup

Requirements:
* python 3.10 or later
* pandas
* openpyxl (for xlsx files)
* odfpy (for ods files)

Suggestions:
* pyarrow

Environment Setup:
```
python3 -m venv env
. env/bin activate
pip install -r requirements.txt
```

## Usage
```
# Example Usage
./xls-parser.py --device-type 'DEVICE' ~/Spreadsheet.xlsx

device,ip

./xls-parser.py --help
usage: xls-parser.py [-h] [--sheet-name SHEET_NAME] [--ip-row-name IP_ROW_NAME] [--hostname-row-name HOSTNAME_ROW_NAME] [--device-type-row-name DEVICE_TYPE_ROW_NAME]
                     [--device-type DEVICE_TYPE [DEVICE_TYPE ...]] [--header-row HEADER_ROW]
                     filename

positional arguments:
  filename              The spreadsheet to open, in xls format

options:
  -h, --help            show this help message and exit
  --sheet-name SHEET_NAME
                        The name of the sheet containing the information to extract
  --ip-row-name IP_ROW_NAME
                        The name of the ip address row within the sheet in order to populate the second column
  --hostname-row-name HOSTNAME_ROW_NAME
                        The name of the hostname row with the sheet in order to populate the first column
  --device-type-row-name DEVICE_TYPE_ROW_NAME
                        The name of the device ID row
  --device-type DEVICE_TYPE [DEVICE_TYPE ...]
                        The device ID to filter for
  --header-row HEADER_ROW
                        The row ID of the column names row, 0 indexed
```
