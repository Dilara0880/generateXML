# Excel to XML Converter

This project includes a Python script that takes an Excel spreadsheet as input, processes the data and generates an XML file based on the data in the Excel file. It also retrieves the exchange rate of USD from the Central Bank of Russia's website on specific dates and calculates and adds a new attribute based on the exchange rate to the XML file.

### Requirements

Python 3.8 or later
pandas library
lxml library
requests library
BeautifulSoup library
logging library

### Instructions

1. Ensure that your Excel file is in the correct format and in the same directory as the Python script. The expected format is provided in ['test_input'](https://github.com/Dilara0880/generateXML/blob/main/test_input.xlsx).
2. Install all the requirements using `python -m pip install -r requirements.txt`. 
3. Use `python script.py` to run the script with Python 3.8 or later.
4. The script will generate an XML file in the same directory, with the data from the Excel file and additional attributes based on the exchange rate of USD.

### Features

- The script can handle Excel files with a specific structure and generates XML files with a specific structure.
- The script retrieves the exchange rate of USD from the [Central Bank of Russia's website](https://www.cbr.ru/currency_base/daily/) on specific dates, calculates and adds a new attribute based on the exchange rate.
- The script includes logging, with timestamps for each log message.
