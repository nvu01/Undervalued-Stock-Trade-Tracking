import glob
import os
import pandas as pd
from io import StringIO
import argparse

# Get the filepath for the latest position statement file
def get_latest_file (folder):
    '''
    Retrieves the latest CSV file in the "position_statement" folder.
    '''
    files = glob.glob(os.path.join(folder, "*"))
    filepath = max(files, key=os.path.getmtime)
    return filepath

# Transform the current position data
def get_current_pos (filepath):
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
            start_indx = i+3
            break

    end_indx = None
    for j in range(start_indx+1, len(lines)):
        if lines[j].strip()=='':
            end_indx = j-1
            break

    # Extract the table
    data = ''.join(lines[start_indx:end_indx])
    df = pd.read_csv(StringIO(data))

    # Transform the dataframe
    df.dropna(subset=['BP Effect'], inplace=True)
    df.reset_index(drop=True, inplace=True)
    df = df[['Instrument','Qty','Trade Price']]
    df.rename(columns={'Instrument':'Symbol'}, inplace=True)

    return df

# Transform new trades data
def get_new_trades (filepath):
    '''
    Processes the new trades data downloaded from TOS.
    Parameter is the filepath to most current CSV file in "account_statement" folder.
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

# Get both previous and new undervalued trades
def filter_new_trades(new_trades, pos_statement_file):
    '''
    Extract the undervalued stock trades from the new trades
    1. Filter out other portfolios' trades for stocks that are also in the "Undervalued" portfolio
    2. Extract buy trades for undervalued stocks based on current positions
    3. Extract closing trades for underavalued stocks based on previous trades
    Return previous trades and new trades
    '''

    # Filter out overlapping stocks
    overlapping = pd.read_excel('overlapping_stocks.xlsx')
    overlapping.drop(columns='Strategy', inplace=True)

    new_trades['Exec Date'] = pd.to_datetime(new_trades['Exec Time'],
                                             format='%m/%d/%y %H:%M:%S').dt.normalize()  # Convert data from 'Exec Time' to datetime64[ns] and extract the dates
    merged = pd.merge(new_trades, overlapping, how='left', on=['Exec Date', 'Symbol', 'Qty', 'Price'], indicator=True)
    new_trades = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge', 'Exec Date'])

    # Split the new trades into buy and sell trades
    buy_to_open = new_trades[(new_trades['Side'] == 'BUY') & (new_trades['Pos Effect'] == 'TO OPEN')]
    sell_to_close = new_trades[(new_trades['Side'] == 'SELL') & (new_trades['Pos Effect'] == 'TO CLOSE')]

    # List of current positions is used to filter buy trades
    current_position = get_current_pos(pos_statement_file)
    undervalued_buy = buy_to_open[buy_to_open['Symbol'].isin(current_position['Symbol'])]

    # List of previous undervalued stock trades is used to filter sell-to-close trade
    previous_trades = pd.read_csv('undervalued_trades.csv')
    filter_stocks = previous_trades['Symbol'].unique()
    undervalued_sell = sell_to_close[sell_to_close['Symbol'].isin(filter_stocks)]
    new_undervalued_trades = pd.concat([undervalued_buy, undervalued_sell])

    return previous_trades, new_undervalued_trades

# Update undervalued trades
def update_undervalued_trades(previous_trades, new_undervalued_trades):
    '''
    Add new trades to previous trades
    Return all undervalued stock trades
    '''

    # Add new records to all undervalued_trades.csv
    rows_to_add = new_undervalued_trades[~new_undervalued_trades['Exec Time'].isin(previous_trades['Exec Time'])]
    trades = pd.concat([previous_trades, rows_to_add])
    trades.reset_index(drop=True, inplace=True)

    return trades

# Process and update undervalued trades
def main(acc_statement_file, pos_statement_file):
    '''
    Apply other functions to update undervalued trades and save all trades to undervalued_trades.csv
    '''
    new_trades = get_new_trades(acc_statement_file)
    previous_trades, new_undervalued_trades = filter_new_trades(new_trades, pos_statement_file)
    trades = update_undervalued_trades(previous_trades, new_undervalued_trades)
    # Save all trades to undervalued_trades.csv
    trades.to_csv('undervalued_trades.csv', index=False)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--pos_statement', help='Optional: A specific file name in position_statement folder')
    parser.add_argument('--acc_statement', help='Optional: A specific file name in account_statement folder')
    args = parser.parse_args()

    # If acc_statement_file is not specified, get_latest_pos_statement will be executed to return the latest file
    if args.acc_statement:
        acc_statement_path = os.path.join('account_statement', args.acc_statement)
    else:
        choice = input("Use latest account statement file? (y/n): ").lower()

        if choice == 'y':
            acc_statement_path = get_latest_file('account_statement')
        else:
            acc_file = input("Enter file name: ").strip()
            acc_statement_path = os.path.join('account_statement', acc_file)

    # If pos_statement_file is not specified, get_latest_pos_statement will be executed to return the latest file
    if args.pos_statement:
        pos_statement_path = os.path.join('position_statement', args.pos_statement)
    else:
        choice = input("Use latest position statement file? (y/n): ").lower()

        if choice == 'y':
            pos_statement_path = get_latest_file('position_statement')
        else:
            pos_file = input("Enter file name: ").strip()
            pos_statement_path = os.path.join('position_statement', pos_file)

    main(acc_statement_path, pos_statement_path)
    
    print(f'Files used: {acc_statement_path}, {pos_statement_path}')