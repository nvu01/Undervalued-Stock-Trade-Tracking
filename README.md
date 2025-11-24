# Undervalued Stock Trade Tracking

This project automates the process of tracking and analyzing undervalued stock trades using a Python ETL pipeline, combined with Power Query and Excel for real-time reporting. The workflow involves accessing the file system to retrieve account and position statements downloaded from Thinkorswim, extracting data from the source CSV files, processing it for undervalued stock trades, and generating an up-to-date portfolio analysis.

## Repository Contents
To protect personal and financial information, the original notebooks and raw CSV files downloaded from TOS are not included in this repository.

- HTML exports of development notebooks: 
    - `first trades.html`
    - `current pos.html`
    - `new trades.html`
- `etl.py`: Production ETL pipeline
- Masked version of `undervalued_trades.csv`
- `Report.xlsx` generated using the masked `undervalued_trades.csv`

## Workflow Overview

Account statements and position statements after Nov 14th 2025 are downloaded every month and include trading data from the 15th of the previous month to the 14th of the current month. These files are processed every month to update the `undervalued_trades.csv` file. This file is then used to generate the Excel file `Report.xlsx`.

## ETL pipeline

The purpose of the ETL pipeline is to retrieve new undervalued stock trades and add them to the `undervalued_trades.csv` file. 

#### Source Files Used During Pipeline Setup:
- The first account statement (**account_statement/2024-11-14-AccountStatement.csv**) include trades from Oct 1st 2024 to Nov 14th 2024. This file was also used to set up the `undervalued_trades.csv` file. Data processing was done in `first trades.ipynb`.
- The second account statement (**account_statement/2025-11-14-AccountStatement.csv**) include all trades from Nov 15th 2024 to Nov 14th 2025. This file was processed in `new trades.ipynb`.
- The first position statement (**position_statement/2025-11-13-PositionStatement.csv**) include the "Undervalued" portfolio's holdings as of Nov 13th 2025. This file was also used to set up the `current_pos.csv` file. Data processing was done in `current pos.ipynb`.

#### ETL Pipeline Functions (etl.py):
- **get_latest_pos_statement** and **get_latest_new_trades**: Functions that fetch the latest position and account statement CSV files from specified folders.
- **get_current_pos**: This function processes the current position data from the "position_statement" folder, specifically extracting the "Undervalued" group, and saves the relevant positions as current_pos.csv.
- **get_new_trades**: This function processes the latest account statement to extract the "Account Trade History" table, filtering relevant buy and sell trades.
- **main**: The core function that updates the undervalued_trades.csv file with new buy-to-open and sell-to-close trades based on the latest position and trade data.

## Excel Reporting

The data in undervalued_trades.csv is used to generate an Excel report that includes:
- Real-time stock prices (using the RTD function in Excel)
- Key performance metrics, including total size, total investment, ROI, open P&L, realized P&L, value weight and monthly P&L.
- A treemap that shows each stock's value weight in the portfolio
- A waterfall chart that visualizes the cumulative profit and loss over time.