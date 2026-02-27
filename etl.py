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
        files = glob.glob(os.path.join(root,'*.csv'))
        for f in files :
            all_files.append(os.path.abspath(f))

    # get total number of files found
    num_files = len(all_files)
    return all_files, num_files

# Retrieve all trade data in a specific period
def get_all_trades (filepath):
    '''
    Process account statements downloaded from TOS.
    Parameter is the filepath to a CSV file in "account_statement" folder.
    1. Load the content of the CSV file.
    2. Find and extract the table named "Account Trade History"
    Return the table as a dataframe
    '''
    with open (filepath) as f:
        lines = f.readlines()

    # Identify the location of the table
    start_indx=None
    for i,line in enumerate(lines):
        if 'Account Trade History' in line:
            start_indx=i+1
            break

    end_indx=None
    for j in range(start_indx+1, len(lines)):
        if lines[j].strip()=='':
            end_indx=j
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

# Use previous trades and position statement to get new trades
def filter_new_trades(previous_trades, all_trades, pos_stmt_file):
    '''
    Extract the undervalued stock trades from the all the new trades
    1. Extract buy trades for undervalued stocks based on current positions
    2. Extract closing trades for underavalued stocks based on previous trades
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
    new_undervalued_trades = pd.concat([undervalued_buy, undervalued_sell])

    # Make sure new rows don't exist in previous_trades
    columns = ['Exec Time', 'Side', 'Pos Effect', 'Symbol', 'Qty', 'Price']
    rows_to_add = pd.merge(new_undervalued_trades, previous_trades, on=columns, how='left', indicator=True)
    rows_to_add = rows_to_add[rows_to_add['_merge'] == 'left_only'].drop(columns='_merge')

    return rows_to_add

def remove_overlapping_stocks(trades):
    '''
    Filter out other portfolios' trades for stocks that are also in the "Undervalued" portfolio
    '''
    # Filter out overlapping stocks
    overlapping = pd.read_excel('overlapping_stocks.xlsx')
    overlapping.drop(columns='Strategy', inplace=True)
    # Convert data from 'Exec Time' to datetime64[ns] and extract the dates
    trades['Exec Date'] = pd.to_datetime(trades['Exec Time'], errors='coerce').dt.normalize()
    merged = pd.merge(trades, overlapping, how='left', on=['Exec Date', 'Symbol', 'Qty', 'Price'], indicator=True)
    filtered_trades = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge', 'Exec Date'])

    return filtered_trades

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

# Process and update undervalued trades
def main():
    '''
    Apply other functions to update undervalued trades and save all trades to undervalued_trades.csv
    1. Get all file paths in "account_statement" and "position_statement"
    2. Use the first file in "account_statement" to generate "undervalued_trades.csv"
    3. Iterates over the remaining account statements, pairing each with the corresponding position statement file.
        - Filters new undervalued trades.
        - Updates the 'undervalued_trades.csv' file with any new undervalued trades.
    '''
    acc_files, acc_num_files = get_files("account_statement")
    print(f'{acc_num_files} files found in "account_statement"')

    pos_files, pos_num_files = get_files("position_statement")
    print(f'{pos_num_files} files found in "position_statement"')

    previous_trades = pd.DataFrame()

    for i in range(len(acc_files)):
        acc_filepath = acc_files[i]
        acc_file_name = os.path.basename(acc_filepath)

        all_trades = get_all_trades(acc_filepath)

        if i == 0:
            new_trades = all_trades.drop(8)
            print(f'First trades retrieved. File used: {acc_file_name}')
        else:
            pos_filepath = pos_files[i - 1]
            pos_file_name = os.path.basename(pos_filepath)

            new_trades = filter_new_trades(previous_trades, all_trades, pos_filepath)
            print(f'Data batch #{i} retrieved. Files used: {acc_file_name}, {pos_file_name}')

        # Update running DataFrame
        previous_trades = pd.concat([previous_trades, new_trades], ignore_index=True)

    all_new_trades = previous_trades
    filtered_trades = remove_overlapping_stocks(all_new_trades)
    updated_trades = update_changes(filtered_trades)

    # Save all trades to undervalued_trades.csv
    updated_trades.to_csv('undervalued_trades.csv', index=False)

    print('Data update complete')

if __name__ == "__main__":
    main()