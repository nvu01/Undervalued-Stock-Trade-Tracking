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

Account statements and position statements after Nov 14th 2025 are downloaded and processed every month to update the `undervalued_trades.csv` file. 
This file is then used to generate the Excel file `Report.xlsm` which shows the portfolio's performance.


## ETL pipeline

The purpose of the ETL pipeline is to process position statements and account statements to extract new undervalued stock trades and add them to the `undervalued_trades.csv` file. 

### Expected Monthly Workflow
1. Monthly data download: Every month, new account statements and position statements are downloaded. 
These files cover trading data from the day after the last date of the previous file up to the current date.
2. Initialize undervalued trades: The first account statement, covering trades from October 1, 2024, to November 14, 2024, is processed.
This initializes the `undervalued_trades.csv` file with the earliest undervalued trades.
3. Iterate over remaining account statements and all position statements:
   - Each subsequent account statement is paired with its corresponding position statement.
   - For each iteration: 
     - New trades are extracted from the account statement.
     - Exclude trades listed in `overlapping_stocks.xlsm`.
     - The open undervalued stock positions are extracted from position statement. 
     - The list of open positions, along with the previous trades from `undervalued_trades.csv`, is used to identify new undervalued trades.
     - Any new undervalued trades are appended to `undervalued_trades.csv`. 
4. Output for reporting: After processing, the updated `undervalued_trades.csv` is used to generate the monthly Excel report (`report.xlsx`).

### ETL Pipeline Functions (etl.py):
- `get_files`: Fetches all CSV files in `account_statement/` or `position_statement/` folders.
- `get_current_pos`: Processes the position statement, specifically extracting the "Undervalued" group, and returns the relevant positions.
- `get_new_trades`: Processes the account statement to extract the "Account Trade History" table and returns all the new trades.
- `filter_new_trades`: Filters the new trades to extract only undervalued stock trades.
- `update_undervalued_trades`: Adds new undervalued trades to the existing table of previous undervalued trades.
- `main`: Iterates through each file in `account_statement/` and `position_statement/`, applies the other functions, and updates `undervalued_trades.csv` with all trades.

### Source Files Used in Development Notebooks:
- The **first account statement** (`account_statement/2024-11-14-AccountStatement.csv`) includes trades from Oct 1st 2024 to Nov 14th 2024. This file is used to generate the `undervalued_trades.csv` file. Data processing was done in `first trades.ipynb` to generate `undervalued_trades.csv` file.
- The **second account statement** (`account_statement/2025-11-14-AccountStatement.csv`) includes all trades from Nov 15th 2024 to Nov 14th 2025. This file was processed in `new trades.ipynb` to append new stock trades to `undervalued_trades.csv`.
- The **first position statement** (`position_statement/2025-11-13-PositionStatement.csv`) includes the "Undervalued" portfolio's holdings as of Nov 13th 2025. This file is used to extract the open positions for undervalued stocks. Data processing was done in `current pos.ipynb`.
- The **Excel file** (`overlapping_stocks.xlsm`) contains a list of trades which are in other portfolios but their tickers are also in the "Undervalued". 
These trades are filtered out of the new trades. This file has to be updated before running `etl.py`.


## Excel Reporting

The data in `undervalued_trades.csv` is used to generate an Excel report (`Report.xlsx`) that includes:
- **Real-time** stock prices (using the RTD function in Excel)
- Key performance metrics: Portfolio's Value, Open and Realized P&Ls, Open and Realized ROIs, Total Spending, Net Spending, Cummulative P&L
- **Treemaps** that show each stock's **open P&L weight** and **investment weight** in the portfolio.
- **Waterfall charts** that visualize the cumulative profit and loss over time.
- Buttons that execute a **VBA macro** script which automates data refresh for all the query tables and pivot tables.