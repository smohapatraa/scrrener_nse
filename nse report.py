import pandas as pd
import os

# Function to calculate RSI(14)
# Directory containing the Excel files
folder_path = 'D:\\nsepy\\dwn_data'

# List to hold the first records
first_records = []

# Loop through each file in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file_name)

        # Read the Excel file
        df = pd.read_excel(file_path)
        delta = df['close'].diff()
        up,down = delta.copy(),delta.copy()
        up[up>0]=0




        # Get the first record and append to the list
        if not df.empty:




            first_record = df.iloc[0].copy()
            first_record['File Name'] = os.path.splitext(file_name)[0]
            sma15 = df['close'].head(15).mean()
            first_record['sma15'] = sma15
            mxml = df['close'].head(3).max()
            first_record['mxml'] = mxml
            mxvl = df['trdqty'].head(3).max()
            first_record['mxvl'] = mxvl
            sma20 = df['close'].head(20).mean()
            first_record['sma20'] = sma20
            sma50 = df['close'].head(50).mean()
            first_record['sma50'] = sma50
            sma100 = df['close'].head(100).mean()
            first_record['sma100'] = sma100
            sma200 = df['close'].head(200).mean()
            first_record['sma200'] = sma200
            first_records.append(first_record)

# Create a DataFrame from the list of first records

result_df = pd.DataFrame(first_records)
result_df['long'] = result_df.apply(lambda row: 'y' if (row['trdqty'] == row['mxvl']) & (row['close'] == row['mxml']) & (row['close'] > row['sma20']) & (row['close'] > row['sma50']) & (row['close'] > row['sma100']) & (row['sma100'] > row['sma200']) else 'n', axis=1)

# Filter positive stocks

filtered_df = result_df[result_df['long'] == 'y']


# Write the result to a new Excel file
output_file_path = 'D:\\nsepy\\nsereport.xlsx'
filtered_df.to_excel(output_file_path, index=False)

print(f'Combined first records saved to {output_file_path}')
