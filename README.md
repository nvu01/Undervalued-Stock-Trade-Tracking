**Disclaimer: This project is intended for educational and informational purposes only. 
The analysis, models, and methodologies used in this project are not intended to serve as financial or investment advice. 
The information presented should not be construed as a recommendation to buy, sell, or hold any securities or assets. 
Investing in the stock market carries inherent risks, and any decisions based on the content of this project are solely the responsibility of the individual. 
Always consult with a qualified financial advisor before making any investment decisions.**

# Undervalued Stock Trade Tracking

This project automates the process of tracking and analyzing undervalued stock trades using a Python ETL pipeline, combined with Power Query and Excel for real-time reporting. The workflow involves accessing the file system to retrieve account and position statements downloaded from Thinkorswim, extracting data from the source CSV files, processing it for undervalued stock trades, and generating an up-to-date portfolio analysis.

## Repository Contents
To protect personal and financial information, the original notebooks, raw CSV files downloaded from TOS, and the outcome `undervalued_trades.csv` are not included in this repository. 
Instead, you can find the HTML versions of the notebooks with some concealed confidential data and a masked version of `undervalued_trades.csv`.

- `etl.py`: Production ETL pipeline
- `masked_undervalued_trades.csv`: Masked version of undervalued trades CSV file
- `Report.xlsm`: A report of portfolio's performance, generated using `masked_undervalued_trades.csv`
- `Terminologies & Formulas.pdf`: A document containing definitions and formulas for metrics used in the report.
- HTML exports of development notebooks: 
    - `first trades.html`
    - `current pos.html`
    - `new trades.html`

## Workflow Overview

Account statements and position statements are downloaded and processed every month to update the `undervalued_trades.csv` file. 
This file is then used to generate the Excel file `Report.xlsm` which shows the portfolio's performance.

## Project Folder and File Structure
```bash
project/
├─ account_statement/
│  ├─ 2024/
│  │  └─ 2024-12-14-AccountStatement.csv       # First account statement to initialize undervalued_trades.csv
│  ├─ 2025/
│  │  └─ 2025-12-14-AccountStatement.csv       # Account statement for new trades in 2025
│  └─ <year>/                                  # Future years
│     ├─ <YYYY-MM-DD>-AccountStatement.csv
│     └─ <YYYY-MM-DD>-AccountStatement.csv     
│
├─ position_statement/
│  ├─ 2025/
│  │  └─ 2025-12-14-PositionStatement.csv      # Open positions for undervalued stocks as of Dec 14th 2025
│  └─ <year>/                                  # Future years
│     ├─ <YYYY-MM-DD>-PositionStatement.csv
│     └─ <YYYY-MM-DD>-PositionStatement.csv
│
├─ etl.py                                       # Main ETL script
├─ undervalued_trades.csv                       # Updated by etl.py
├─ Overlapping Stocks.xlsx                      # Filter trades present in other portfolios
├─ Symbol Change.xlsx                           # Stores ticker changes and trade updates
└─ Report.xlsm                                  # Imports undervalued_trades.csv to generate performance reports
```
Unlike other later years, the year 2025 only contains one account statement and one position statement because:
- This project started at the end of 2025, so monthly snapshots of positions were not available. 
- All trades in 2025 were buy transactions, so there were no sell transactions to track. 

## ETL pipeline

The purpose of the ETL pipeline is to process position statements and account statements to extract new undervalued stock trades and add them to the `undervalued_trades.csv` file. 

### Expected Monthly Workflow
1. Monthly data download: Every month, new account statements and position statements are downloaded and stored in the corresponding trading year folder (`account_statement/<year>/` or `position_statement/<year>/`). 
These files cover trading data from the day after the last date of the previous file up to the current date.
2. Initialize undervalued trades: The first account statement, covering trades from October 1, 2024, to December 14, 2024, is processed.
This initializes the `undervalued_trades.csv` file with the earliest undervalued trades in 2024.
3. Iterate over the account statements and position statements in `account_statement/<year>/` and `position_statement/<year>/`:
   - Each subsequent account statement is paired with its corresponding position statement.
   - For each iteration: 
     - New trades are extracted from the account statement.
     - The current open positions for undervalued stocks are extracted from position statement. 
     - Previous trades from `undervalued_trades.csv` are combined with the current positions to identify new undervalued trades.
     - Changes in ticker symbols or trade data are applied using `Symbol Change.xlsx`.
     - Exclude trades listed in `Overlapping Stocks.xlsm`.
     - Any new undervalued trades are appended to `undervalued_trades.csv`. 
4. Output for reporting: After processing, the updated `undervalued_trades.csv` is used to generate the monthly Excel report (`report.xlsx`).

Notes on trade data processing in 2025: Only one account statement and one position statement exist because the project started at the end of 2025. 
All trades in 2025 were buy transactions, so the final position snapshot is sufficient to identify undervalued stock trades. 
There is no need to update `undervalued_trades.csv` monthly for 2025, since sell transactions do not affect the identification of undervalued trades.

### ETL Pipeline Functions (etl.py):
- `get_files`: Fetches all CSV files in `account_statement/` or `position_statement/` folders.
- `get_current_pos`: Processes the position statement, specifically extracting the "Undervalued" group, and returns the relevant positions.
- `get_new_trades`: Processes the account statement to extract the "Account Trade History" table and returns all the new trades.
- `filter_new_trades`: Filters the new trades to extract only undervalued stock trades.
- `update_changes`: Update the new trades with any changes in symbol or trade data as recorded in `Symbol Change.xlsx`.
- `remove_overlapping_stocks`: Excludes trades that belong to other portfolios but have the same ticker as an undervalued stock, based on `Overlapping Stocks.xlsx`.
- `main`: Iterates through each file in a given year, applies all functions above, and updates `undervalued_trades.csv`.

### Source Files Used in Development Notebooks:
- The **first account statement** (`account_statement/2024/2024-12-14-AccountStatement.csv`) includes trades from Oct 1st 2024 to Dec 14th 2024. This file is used to generate the `undervalued_trades.csv` file. Data processing was done in `first trades.ipynb` to generate `undervalued_trades.csv` file.
- The **second account statement** (`account_statement/2025/2025-12-14-AccountStatement.csv`) includes all trades from Dec 15th 2024 to Nov 14th 2025. This file was processed in `new trades.ipynb` to append new stock trades to `undervalued_trades.csv`.
- The **first position statement** (`position_statement/2025/2025-12-143-PositionStatement.csv`) includes the "Undervalued" portfolio's holdings as of Dec 14th 2025. This file is used to extract the open positions for undervalued stocks. Data processing was done in `current pos.ipynb`.
- The **Excel file** (`Overlapping Stocks.xlsx`) contains a list of trades which are in other portfolios but their tickers are also in the "Undervalued". 
These trades are filtered out of the new trades. This file has to be updated before running `etl.py`.
- Excel file (`Symbol Change.xlsx`) stores changes in ticker symbol and trade data due to rebranding, merger or acquisation. 
This file has to be manually updated before running etl.py.

## Excel Reporting

The data in `undervalued_trades.csv` is used to generate an Excel report (`Report.xlsx`) that includes:
- **Real-time** stock prices (using the RTD function in Excel)
- Key performance metrics: Portfolio's Value, Open and Realized P&Ls, Open and Realized ROIs, Total Spending, Net Spending, Cummulative P&L
- **Treemaps** that show each stock's **open P&L weight** and **investment weight** in the portfolio.
- **Waterfall charts** that visualize the cumulative profit and loss over time.
- Buttons that execute a **VBA macro** script which automates data refresh for all the query tables and pivot tables.