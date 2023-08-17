# XML to Excel Data Extraction Manual

This manual outlines the usage of the provided Python script to extract financial data from a web API using XML configuration and save it into an Excel file. The script uses libraries such as `xml.etree.ElementTree`, `requests`, `time`, and `openpyxl`.

## Table of Contents
1. [Introduction](#introduction)
2. [Installation](#installation)
3. [Usage](#usage)
4. [Configuration](#configuration)
5. [Output](#output)
6. [Limitations](#limitations)

## 1. Introduction <a name="introduction"></a>
The provided Python script reads configuration data from an XML file, interacts with a web API to fetch financial data for a specific date range, and then saves the extracted data into an Excel spreadsheet. This can be useful for financial data analysis and record-keeping.

## 2. Installation <a name="installation"></a>
Before running the script, make sure you have the necessary packages installed. You can install them using the following commands:

```bash
pip install requests openpyxl
```

## 3. Usage <a name="usage"></a>
To use the script, follow these steps:

1. Place the script (`script.py`) and the XML configuration file (`setting2.xml`) in the same directory.

2. Open a terminal or command prompt and navigate to the directory containing the script and XML file.

3. Run the script using the following command:

```bash
python script.py
```

4. The script will execute, fetching financial data from the specified web API based on the configuration provided in the XML file.

## 4. Configuration <a name="configuration"></a>
Before running the script, you need to configure the `setting2.xml` file with the relevant information. The XML file contains the following configuration parameters:

- `<startyear>`: Start year for the data extraction.
- `<startmonth>`: Start month for the data extraction.
- `<endyear>`: End year for the data extraction.
- `<endmonth>`: End month for the data extraction.
- `<url>`: URL of the web API to fetch data from.
- `<stockNo>`: Stock number for the financial data.
- `<excelname>`: Name of the Excel file to be generated.

Modify these values in the XML file to customize the data extraction process.

## 5. Output <a name="output"></a>
The script will generate an Excel file named as specified in the `<excelname>` parameter. The spreadsheet will have a sheet named "Excel Everyday records" containing extracted financial data. The data will be organized in columns as follows:

- Date
- Profit Margin
- Closing Price
- Starting Price
- Highest Point
- Lowest Point
- Transaction Amount
- Symphony of Transaction

## 6. Limitations <a name="limitations"></a>
- The script is dependent on the structure and availability of the web API for fetching data. If the API structure changes, the script might need modifications.
- The script may not handle all possible error scenarios gracefully.
- The script uses a time delay of 3 seconds (`time.sleep(3)`) between API requests. This delay might need adjustment based on the API's rate limits and your specific use case.

**Note:** This manual is provided as a general guide. Be sure to review and understand the script before running it, especially if it will be handling sensitive data or interacting with critical systems.
