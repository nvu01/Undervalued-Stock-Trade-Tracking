import glob
import os
import pandas as pd
from io import StringIO


# Get all account statements or position statements
def get_files(folder):
    '''
    Retrieve all files in a folder
    Return a list of file paths
    Return the number files in the folder
    '''
    all_files = []
    for root, dirs, files in os.walk(folder):
        files = glob.glob(os.path.join(root, '*.csv'))
        for f in files:
            all_files.append(os.path.abspath(f))

    # Sort files in ascending order (oldest first, newest last)
    all_files.sort()

    # get total number of files found
    num_files = len(all_files)
    return all_files, num_files

# Retrieve all trade data in a specific period
def get_all_trades(filepath):
    '''
    Process account statements downloaded from TOS.
    Parameter is the filepath to a CSV file in "account_statement" folder.
    1. Load the content of the CSV file.
    2. Find and extract the table named "Account Trade History"
    Return the table as a dataframe
    '''
    with open(filepath) as f:
        lines = f.readlines()

    # Identify the location of the table
    start_indx = None
    for i, line in enumerate(lines):
        if 'Account Trade History' in line:
            start_indx = i + 1
            break

    end_indx = None
    for j in range(start_indx + 1, len(lines)):
        if lines[j].strip() == '':
            end_indx = j
            break

    # Extract the table
    data = ''.join(lines[start_indx:end_indx])
    df = pd.read_csv(StringIO(data))

    # Keep only the important columns
    df = df[['Exec Time', 'Side', 'Pos Effect', 'Symbol', 'Qty', 'Price']]

    return df


# Retrieve and transform the current position data
def get_current_pos(filepath):
    '''
    Processes the current position data downloaded from TOS.
    Parameter is the filepath to the most current CSV file in "position_statement" folder.
    1. Load the content of the CSV file.
    2. Find and extract the table named "Undervalued"
    3. Save the table as "current_pos.csv" in the parent folder
    Return the table as a dataframe
    '''
    with open(filepath) as f:
        lines = f.readlines()

    # Identify the location of the table
    start_indx = None
    for i, line in enumerate(lines):
        if 'Group "Undervalued"' in line:
            start_indx = i + 3
            break

    end_indx = None
    for j in range(start_indx + 1, len(lines)):
        if lines[j].strip() == '':
            end_indx = j - 1
            break

    # Extract the table
    data = ''.join(lines[start_indx:end_indx])
    df = pd.read_csv(StringIO(data))

    # Transform the dataframe
    df.dropna(subset=['BP Effect'], inplace=True)
    df.reset_index(drop=True, inplace=True)
    df = df[['Instrument', 'Qty', 'Trade Price']]
    df.rename(columns={'Instrument': 'Symbol'}, inplace=True)

    return df

def update_changes(trades):
    '''
    Update changes in ticker symbol
    1. Get the new changes from "Symbol Change.xlsx"
    2. Remove records of old symbols
    3. Add new changes to the trade data
    '''
    # Import the Excel file that stores changes
    ticker_change = pd.read_excel('Symbol Change.xlsx')
    # Find new changes that are not yet updated in trades
    replace_rows = ticker_change.loc[ticker_change['Old Symbol'].isin(trades['Symbol']), ['Exec Time', 'Side', 'Pos Effect', 'Symbol', 'Qty','Price']]
    # Remove the records of old symbols
    trades = trades[~trades['Symbol'].isin(ticker_change['Old Symbol'])]
    # Add new updated records to trades
    updated_trades = pd.concat([trades, replace_rows])

    return updated_trades

# Use previous trades and position statement to get new trades
def filter_new_trades(previous_trades, all_trades, pos_stmt_file):
    '''
    Extract the undervalued stock trades from the all the new trades
    1. Extract buy trades for undervalued stocks based on current positions
    2. Extract closing trades for underavalued stocks based on previous trades
    3. Update any change in symbol or trade data
    Return new undervalued trades
    '''
    # Split the new trades into buy and sell trades
    buy_to_open = all_trades[(all_trades['Side'] == 'BUY') & (all_trades['Pos Effect'] == 'TO OPEN')]
    sell_to_close = all_trades[(all_trades['Side'] == 'SELL') & (all_trades['Pos Effect'] == 'TO CLOSE')]

    # List of current positions is used to filter buy trades
    current_position = get_current_pos(pos_stmt_file)
    undervalued_buy = buy_to_open[buy_to_open['Symbol'].isin(current_position['Symbol'])]

    # List of previous undervalued stock trades is used to filter sell-to-close trade
    filter_stocks = previous_trades['Symbol'].unique()
    undervalued_sell = sell_to_close[sell_to_close['Symbol'].isin(filter_stocks)]

    # Combine undervalued_buy and undervalued_sell
    undervalued_trades = pd.concat([undervalued_buy, undervalued_sell])

    # Update any symbol changes
    updated_undervalued_trades = update_changes(undervalued_trades)

    # Convert "Exec Time" in updated_undervalued_trades before comparing it with previous_trades
    updated_undervalued_trades['Exec Time'] = pd.to_datetime(updated_undervalued_trades['Exec Time'], format='%m/%d/%y %H:%M:%S')
    # Exclude trades already existing in previous_trades
    columns = ['Exec Time', 'Side', 'Pos Effect', 'Symbol', 'Qty', 'Price']
    rows_to_add = pd.merge(updated_undervalued_trades, previous_trades, on=columns, how='left', indicator=True)
    rows_to_add = rows_to_add[rows_to_add['_merge'] == 'left_only'].drop(columns='_merge')

    return rows_to_add

def remove_overlapping_stocks(trades):
    '''
    Filter out trades that belong to other portfolios but have the same ticker as an undervalued stock
    '''
    # Filter out overlapping stocks
    overlapping = pd.read_excel('Overlapping Stocks.xlsx')
    overlapping.drop(columns='Strategy', inplace=True)
    # Extract the dates from 'Exec Time' to compare with 'Exec Date' in overlapping
    trades['Exec Date'] = trades['Exec Time'].dt.normalize()
    merged = pd.merge(trades, overlapping, how='left', on=['Exec Date', 'Symbol', 'Qty', 'Price'], indicator=True)
    filtered_trades = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge', 'Exec Date'])

    return filtered_trades


# Process and update undervalued trades
def main(year):
    '''
    Apply other functions to update undervalued trades and save all trades to undervalued_trades.csv
    1. Get all file paths in the input year folder in "account_statement" and "position_statement"
    2. If the year is 2024, reinitialize undervalued_trades.csv
    3. Iterates over the account statements, pairing each with the corresponding position statement file.
        - Filters new undervalued trades.
        - Updates the 'undervalued_trades.csv' file with any new undervalued trades.
    '''
    acc_folder = os.path.join('account_statement', year)
    acc_files, acc_num_files = get_files(acc_folder)
    print(f'{acc_num_files} file(s) found in "account_statement/{year}"')

    if year == '2024':
        acc_filepath = acc_files[0]
        first_trades = get_all_trades(acc_filepath)
        stocks = ['YRD', 'XYF', 'CCRN', 'LX', 'SITC', 'AMTD', 'GSL', 'HG', 'PDS']
        filtered_trades = first_trades.loc[first_trades['Symbol'].isin(stocks), :].copy()
        filtered_trades['Exec Time'] = pd.to_datetime(filtered_trades['Exec Time'], format='%m/%d/%y %H:%M:%S')
        print('Processed trades in 2024')
    else:
        pos_folder = os.path.join('position_statement', year)
        pos_files, pos_num_files = get_files(pos_folder)
        print(f'{pos_num_files} file(s) found in "position_statement/{year}"')

        previous_trades = pd.read_csv('undervalued_trades.csv')
        previous_trades['Exec Time'] = pd.to_datetime(previous_trades['Exec Time'], format='%Y-%m-%d %H:%M:%S')

        for i in range(len(acc_files)):
            acc_filepath = acc_files[i]
            pos_filepath = pos_files[i]

            all_trades = get_all_trades(acc_filepath)
            new_trades = filter_new_trades(previous_trades, all_trades, pos_filepath)

            # Update running DataFrame
            previous_trades = pd.concat([previous_trades, new_trades], ignore_index=True)

            acc_file_name = os.path.basename(acc_filepath)
            pos_file_name = os.path.basename(pos_filepath)
            print(f'Files used: {acc_file_name}, {pos_file_name}')

        filtered_trades = remove_overlapping_stocks(previous_trades)

    # Save all trades to undervalued_trades.csv
    filtered_trades.to_csv('undervalued_trades.csv', index=False)

    print('Data update complete')

if __name__ == "__main__":
    print('Enter the trading year: ')

    trading_year = input()
    main(trading_year)