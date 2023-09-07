import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

from sklearn.ensemble import IsolationForest
from sklearn.impute import SimpleImputer
from sklearn.ensemble import IsolationForest
import pandas as pd
from datetime import datetime, timedelta
import requests
import seaborn as sns
import matplotlib.pyplot as plt
import io
from openpyxl import Workbook, load_workbook
import time
import os
import sys
from tkinter.scrolledtext import ScrolledText

start_time = time.time()

def load_data(file_path):
    data = pd.read_excel(file_path)
    replicate_print(f'The dataset was loaded...')
    return data

def calculate_time_periods(data):
    data['First deposit date'] = pd.to_datetime(data['First deposit date'])
    
    last_day = data['First deposit date'].max().date()

    def replace_time_period(date):
        if pd.notna(date):
            time_passed = (last_day - date.date()).days
            
            if 0 < time_passed <= 7:
                start_date = last_day - timedelta(days=7)
                return f"[{start_date} - {last_day - timedelta(days=1)}]"
            elif 7 < time_passed <= 14:
                start_date = last_day - timedelta(days=14)
                return f"[{start_date} - {last_day - timedelta(days=8)}]"
            elif 28 < time_passed <= 35:
                start_date = last_day - timedelta(days=35)
                return f"[{start_date} - {last_day - timedelta(days=29)}]"
            elif 63 < time_passed <= 70:
                start_date = last_day - timedelta(days=70)
                return f"[{start_date} - {last_day - timedelta(days=64)}]"
        else:
            return 'Missing'

    data['Time Period'] = data['First deposit date'].apply(replace_time_period)
    replicate_print(f'Time period were detected...')

    return data

def calculate_deal_type(row):
    cpa_value = row['CPA']
    rs_value = row['RS']

    if cpa_value != 0 and rs_value == 0:
        return 'CPA'
    elif rs_value != 0 and cpa_value == 0:
        return 'RS'
    elif (rs_value > 0 or rs_value < 0) and cpa_value != 0:
        return 'CPA+RS'    
    else:
        return 'None'

def preprocess_data(data):
    data = calculate_time_periods(data)

    data = data[['Deposits count', 'Deposit amount', 'Bets amount', 'Company profit (total)',
                'RS', 'CPA', 'Commission amount', 'Bonus amount', 'Affiliate ID', 'Time Period', 'Player ID', 'Country']]
    data = data.groupby(['Country','Affiliate ID', 'Time Period']).agg({
        'Deposits count': 'sum',
        'Deposit amount': 'sum',
        'Bets amount': 'sum',
        'Company profit (total)': 'sum',
        'RS': 'sum',
        'CPA': 'sum',
        'Commission amount': 'sum',
        'Bonus amount': 'sum',
        'Player ID': 'count'  # Count the occurrences of Player IDs
    }).reset_index()
    
    data['Deal Type'] = data.apply(calculate_deal_type, axis=1)
    data.reset_index(inplace=True)
    data = data.groupby(['Country','Affiliate ID', 'Time Period', 'Deal Type']).agg({
        'Deposits count': 'sum',
        'Deposit amount': 'sum',
        'Bets amount': 'sum',
        'Company profit (total)': 'sum',
        'Commission amount': 'sum',
        'Bonus amount': 'sum',
        'Player ID': 'sum'  # Count the occurrences of Player IDs
    }).reset_index()
    data = data[data['Player ID']>5]
    replicate_print(f'The dataset was processed')
    return data

def apply_isolation_forest_per_country(data):
    unique_countries = data['Country'].unique()

    for country in unique_countries:
        country_data = data[data['Country'] == country]
        
        # Drop unnecessary columns
        country_data = country_data.drop(['Country', 'Affiliate ID', 'Time Period', 'Deal Type'], axis=1)
        
        # Handle missing values (NaNs) by imputing with median values
        imputer = SimpleImputer(strategy='median')
        country_data_imputed = imputer.fit_transform(country_data)
        
        # Apply Isolation Forest
        model = IsolationForest(contamination='auto', random_state=42)
        model.fit(country_data_imputed)
        
        # Get anomaly scores and add them to the results
        anomaly_scores = model.decision_function(country_data_imputed)
        
        # Merge the anomaly scores back to the original data for this country
        data.loc[data['Country'] == country, 'Anomaly Score'] = anomaly_scores
    
    return data


def scale_anomaly_scores(data):
    # Ensure 'Anomaly Score' is of a numeric data type (e.g., float)
    data['Anomaly Score'] = pd.to_numeric(data['Anomaly Score'], errors='coerce')
    
    # Check if there are any NaN values in 'Anomaly Score' after conversion
    if data['Anomaly Score'].isnull().any():
        replicate_print("Warning: Some 'Anomaly Score' values couldn't be converted to numeric.")

    min_score = data['Anomaly Score'].min()
    max_score = data['Anomaly Score'].max()

    data['Scaled Anomaly Score'] = (data['Anomaly Score'] - min_score) / (max_score - min_score)

    return data

def browse_input_file():
    replicate_print(f'Browsing input file...')
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)

def browse_output_dir():
    replicate_print(f'Browsing output file...')
    output_dir = filedialog.askdirectory()
    if output_dir:
        entry_output_dir.delete(0, tk.END)
        entry_output_dir.insert(0, output_dir)

def replicate_print(message):
    current_time = datetime.now().strftime("%H:%M:%S.%f")[:-3]  # Include milliseconds
    message_with_timestamp = f"[{current_time}] {message}"
    
    console_text.config(state=tk.NORMAL)  # Enable the secondary text widget
    console_text.insert(tk.END, f"{message_with_timestamp}\n")

    start_index = 1.0
    while True:
        start_index = console_text.search("[", start_index, stopindex=tk.END)
        if not start_index:
            break
        end_index = console_text.search("]", start_index)
        if not end_index:
            break
        console_text.tag_add("date_tag", start_index, end_index + " +1c")
        console_text.tag_config("date_tag", foreground="green")
        start_index = end_index + " +1c"

    if message == "Analysis completed, and results saved to separate Excel sheets by 'Country'.":
        end_text = "\n===END===\n"
        console_text.insert(tk.END, end_text, "center_tag")
        console_text.tag_configure("center_tag", justify="center")

    console_text.config(state=tk.DISABLED)  # Disable the secondary text widget
    console_text.see(tk.END)  # Scroll to the end

# Redirect stdout to replicate prints to the secondary console
sys.stdout = replicate_print


def run_analysis():
    replicate_print(f'Running analysis to detect anomalies...')
    input_file_path = entry_file_path.get()
    output_dir = entry_output_dir.get()

    if not input_file_path:
        messagebox.showerror("Error", "Please choose an input file.")
        return

    if not output_dir:
        messagebox.showerror("Error", "Please choose an output directory.")
        return

    # Check if the input file exists
    if not os.path.exists(input_file_path):
        messagebox.showerror("Error", "Input file does not exist.")
        return

    # Check if the output directory exists
    if not os.path.exists(output_dir):
        messagebox.showerror("Error", "Output directory does not exist.")
        return

    # Set the path for the output Excel file
    output_file_path = os.path.join(output_dir, "anomalies_checker_dataset.xlsx")

    data = load_data(input_file_path)
    data = preprocess_data(data)
    #data = apply_isolation_forest_per_country(data)
    #data = scale_anomaly_scores(data)

    try:
        replicate_print('Alomst there. It may take some time...')
        #data = scale_anomaly_scores(data)
        grouped = data.groupby('Country')
        output_file_path = os.path.join(output_dir, "anomalies_checker_dataset.xlsx")
        with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
            for country, country_data in grouped:
                # Apply Isolation Forest to the country's data
                replicate_print(f'Analysing: {country}')
                country_data = apply_isolation_forest_per_country(country_data)
                country_data = scale_anomaly_scores(country_data)
                replicate_print(f'Scaling was applied to: {country}')
                # Save the country's data with anomaly scores to a separate sheet
                country_data.to_excel(writer, sheet_name=country, index=False)
        
        messagebox.showinfo("Success", "Analysis completed, and results saved to separate Excel sheets by 'Country'.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

    replicate_print("Analysis completed, and results saved to separate Excel sheets by 'Country'.")

root = tk.Tk()
root.title("Anomaly Detection GUI")

# Create and place widgets on the window
label_file = tk.Label(root, text="Select Input File:")
label_file.grid(row=0, column=0, padx=10, pady=5)
entry_file_path = tk.Entry(root, width=40)
entry_file_path.grid(row=0, column=1, padx=10, pady=5)
button_browse_file = tk.Button(root, text="Browse", command=browse_input_file)
button_browse_file.grid(row=0, column=2, padx=10, pady=5)

label_output_dir = tk.Label(root, text="Select Output Directory:")
label_output_dir.grid(row=1, column=0, padx=10, pady=5)
entry_output_dir = tk.Entry(root, width=40)
entry_output_dir.grid(row=1, column=1, padx=10, pady=5)
button_browse_output_dir = tk.Button(root, text="Browse", command=browse_output_dir)
button_browse_output_dir.grid(row=1, column=2, padx=10, pady=5)

button_run = tk.Button(root, text="Run Analysis", command=run_analysis)
button_run.grid(row=2, column=1, pady=10)

console_text = ScrolledText(root, wrap=tk.WORD, height=10, width=50)
console_text.grid(row=3, column=0, columnspan=3, padx=10, pady=10)
console_text.config(state=tk.DISABLED)

class ReplicatedConsoleRedirector:
    def __init__(self, callback):
        self.callback = callback

    def write(self, message):
        self.callback(message)

    def flush(self):
        pass

# Redirect stdout to replicate prints to the secondary console
sys.stdout = ReplicatedConsoleRedirector(replicate_print)

# Run the GUI application
root.mainloop()
