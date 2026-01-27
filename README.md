# Undervalued Stock Trade Tracking

This project automates the process of tracking and analyzing undervalued stock trades using a Python ETL pipeline, combined with Power Query and Excel for real-time reporting. The workflow involves accessing the file system to retrieve account and position statements downloaded from Thinkorswim, extracting data from the source CSV files, processing it for undervalued stock trades, and generating an up-to-date portfolio analysis.

## Repository Contents
To protect personal and financial information, the original notebooks, raw CSV files downloaded from TOS, and the outcome `undervalued_trades.csv` are not included in this repository. 
Instead, you can find the HTML versions of the notebooks with some concealed confidential data and a masked version of `undervalued_trades.csv`.

- HTML exports of development notebooks: 
    - `first trades.html`
    - `current pos.html`
    - `new trades.html`
- `etl.py`: Production ETL pipeline
- `masked_undervalued_trades.csv`: Masked version of undervalued trades CSV file
- `Report.xlsm`: A report of portfolio's performance, generated using `masked_undervalued_trades.csv`


## Workflow Overview

Account statements and position statements after Nov 14th 2025 are downloaded and processed every month to update the `undervalued_trades.csv` file. This file is then used to generate the Excel file `Report.xlsm` which shows the portfolio's performance.


## ETL pipeline

The purpose of the ETL pipeline is to process current position statement and account statement to extract new undervalued stock trades and add them to the `undervalued_trades.csv` file. 

#### ETL Pipeline Functions (etl.py):
- `get_latest_pos_statement` and `get_latest_new_trades`: Functions that fetch the latest position and account statement CSV files from specified folders.
- `get_current_pos`: This function processes the current position data from the "position_statement" folder, specifically extracting the "Undervalued" group, and returns the relevant positions.
- `get_new_trades`: This function processes the latest account statement from the "account_statement" to extract the "Account Trade History" table and return all the new trades.
- `main`: The core function that updates the `undervalued_trades.csv` file with new trades based on the latest position and new trade data.

#### Source Files Used During Pipeline Setup:
- The **first account statement** (`account_statement/2024-11-14-AccountStatement.csv`) include trades from Oct 1st 2024 to Nov 14th 2024. This file was used to set up the table containing the first undervalued trades. Data processing was done in `first trades.ipynb` to generate `undervalued_trades.csv` file.
- The **second account statement** (`account_statement/2025-11-14-AccountStatement.csv`) include all trades from Nov 15th 2024 to Nov 14th 2025. This file was processed in `new trades.ipynb` to append new stock trades to `undervalued_trades.csv`.
- The **first position statement** (`position_statement/2025-11-13-PositionStatement.csv`) include the "Undervalued" portfolio's holdings as of Nov 13th 2025. This file was processed in`current pos.ipynb` to set up the `current_pos.csv` file which is used in `new trades.ipynb` to filter undervalued stock trades from all new trades.
- The **Excel file** (`overlapping_stocks.xlsm`) contains a list of trades which are in other portfolios but their tickers are also in the "Undervalued". These trades will be filtered out of the new trades. This file has to be updated before running "etl.py".


## Excel Reporting

The data in undervalued_trades.csv is used to generate an Excel report that includes:
- **Real-time** stock prices (using the RTD function in Excel)
- Key performance metrics: portfolio's total size, total investment, ROI, open P&L, realized P&L, value weight and monthly P&L.
- **Treemaps** that show each stock's P&L weight and investment weight in the portfolio.
- **Waterfall charts** that visualize the cumulative profit and loss over time.
- A button that executes a **VBA macro** script which automates data refresh for all the query tables and pivot tables in the right order.