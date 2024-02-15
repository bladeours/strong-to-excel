# Strong To Excel

Script to create excel sheet from export data from [Strong app](https://www.strong.app).

## Installation

Use `requirements.txt` to install necessary modules.

```bash
pip install -r requirements.txt
```

## How it looks?

![example](example.png)

## Export Data From Strong

Here is official tutorial how to do it: [link](https://help.strongapp.io/article/235-export-workout-data).

## Help

```bash
usage: strong-to-excel [-h] [-l {DEBUG,INFO,WARNING,ERROR,CRITICAL}] [-o OUTPUT] [-i INPUT]

program to import data from Strong app to Excel sheet

options:
  -h, --help            show this help message and exit
  -l {DEBUG,INFO,WARNING,ERROR,CRITICAL}, --logging {DEBUG,INFO,WARNING,ERROR,CRITICAL}
                        do no print to std output
  -o OUTPUT, --output OUTPUT
                        output file name, default - strong-<timestamp>.xlsx
  -i INPUT, --input INPUT
                        input file name, default - strong.csv
```

## Usage

```bash
# creates file strong.xlsx from file strong3231.csv
python3 strong-to-excel.py -i strong3231.csv -o strong.xlsx
```
