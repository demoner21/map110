import os
import re
import csv
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox

def find_pattern_below(content, result_matrix):
    pattern_below = r'\(\d{4}\)\(\d{4}\)\(\d{4}\)\(\d{4}\)\(\d{4}\)\(\d{4}\)'
    matches = re.findall(pattern_below, content)

    if matches:
        num_columns = len(matches[0].strip('()').split(')('))
        result_matrix.extend([[] for _ in range(num_columns)])

        for i, match in enumerate(matches):
            # Split the match into individual elements and convert them to integers
            elements = [int(x) for x in match.strip('()').split(')(')]

            # Update the result matrix by adding the elements to the respective columns
            for j, element in enumerate(elements):
                result_matrix[j].append(element)

def update_data_value(data_value):
    # Update the data value here based on your requirements
    # Increment the data_value by 1
    return int(data_value) + 1

def process_xml_file(file_name, column_index):
    # Get the file name without extension
    file_name_without_extension = os.path.splitext(file_name)[0]

    # Read the XML file
    with open(file_name, 'r') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'xml')

    pattern = r'P\.01\(\d+\)\(\d+\)\(\d+\)\(\d+\)\(1-1:\d+\.\d+\.\d+\)\(kW\)\(1-1:\d+\.\d+\.\d+\)\(kvar\)\(1-1:\d+\.\d+\.\d+\)\(kvar\)\(1-1:\d+\.\d+\.\d+\)\(kW\)\(1-1:\d+\.\d+\.\d+\)\(kvar\)\(1-1:\d+\.\d+\.\d+\)\(kvar\)'
    matches = soup.find_all(string=re.compile(pattern))
    data_pattern = r'P\.01\((\d+)\)'

    if matches:
        num_columns = len(matches)

        result_matrix = [[] for _ in range(num_columns)]

        for match in matches:
            find_pattern_below(match, result_matrix)

        # Extract the specified column values
        column_values = result_matrix[column_index]

        # Starting timestamp
        timestamp = datetime.strptime('00:00:00', '%H:%M:%S')

        clean_data = []  # List to store the clean data

        # Append the clean data to the list
        def append_clean_data(data_value, timestamp_str, value, kWh):
            # Skip appending the data if the current timestamp is "23:45:00" and matches the previous timestamp
            if timestamp_str == '23:45:00' and clean_data[-1]['Hora'] == timestamp_str:
                return

            kWh = value * 0.25
            clean_data.append({"Data": data_value, "Hora": timestamp_str, "kW": value, "kWh": kWh})

        data_match = None  # Initialize data_match
        data_value = None  # Initialize data_value
        cycle_count = 0  # Track the number of completed cycles
        intervals_per_cycle = 24 * 60 // 15  # Calculate the number of 15-minute intervals in a 24-hour cycle

        daily_sums = []  # List to store daily sums
        current_day_sum = 0  # Initialize current day sum
        current_day = None  # Initialize current day

        for value in column_values:
            # Add timestamp to each value in the specified column
            timestamp_str = timestamp.strftime('%H:%M:%S')
            timestamp += timedelta(minutes=15)
            kWh = value * 0.25

            # Find the corresponding date pattern and update the data value if it's different from the previous match
            new_data_match = re.search(data_pattern, content)
            if new_data_match and new_data_match.group(1) != data_match:
                data_match = new_data_match.group(1)
                data_value = data_match[1:-6]  # Remove the first character and last 6 characters

            # Append the clean data to the list
            append_clean_data(data_value, timestamp_str, value, kWh)

            # Check if a 24-hour cycle is completed
            if timestamp_str == '00:00:00':
                cycle_count += 1
                if cycle_count % intervals_per_cycle == 0:
                    data_value = update_data_value(data_value)
                    append_clean_data(data_value, timestamp_str, value, kWh)

            # Check if the timestamp is "23:45:00" and print the updated data value
            if timestamp_str == '23:45:00':
                if data_value:
                    data_value = update_data_value(data_value)
                    append_clean_data(data_value, timestamp_str, value, kWh)

            # Calculate daily sum
            current_day = timestamp.date()
            current_day_sum += kWh

            # Check if a new day has started
            if timestamp_str == '23:45:00':
                daily_sums.append(current_day_sum)
                current_day_sum = 0

        # Sum the values in the specified column
        column_sum = sum(column_values)

        # Multiply the column sum by 0.25
        multiplied_sum = column_sum * 0.25

        total_sum = sum(daily_sums) * 0.25

        total_sum_data = {"Data": "Total Sum", "Hora": "", "kW": column_sum, "kWh": multiplied_sum}

        clean_data.append(total_sum_data)

        column_names = [
            "Columns - 1 - kW", "Columns - 2 kvar", "Columns - 3 kvar", 
            "Columns - 4 kWh", "Columns - 5 kvar", "Columns - 6 kvar"
        ]
        column_name = column_names[column_index]

        # Add time values to the plot suptitle
        time_from, time_until = extract_time_values_from_xml(file_name)

        # Create plot
        plt.plot(range(len(daily_sums)), daily_sums, color='black')
        plt.xlabel('Day')
        plt.ylabel('{} (kWh)'.format(multiplied_sum))
        plt.suptitle('Sum of {} per Day'.format(column_name))
        plt.title('Date Started: {} - Until: {}'.format(time_from, time_until))
        plt.xticks(range(len(daily_sums)), [f'Day {i+1}' for i in range(len(daily_sums))], rotation=45)

        # Save the plot to an image file
        plot_file_path = f"{file_name_without_extension}.png"
        plt.savefig(plot_file_path)
        plt.clf()  # Clear the plot to release memory
        plt.close()  # Close the plot

        # Save the clean data to an Excel file
        excel_file_path = f"{file_name_without_extension}.xlsx"
        save_clean_data(clean_data, excel_file_path)

        # Save the clean data to a CSV file
        csv_file_path = f"{file_name_without_extension}.csv"
        save_clean_data_to_csv(clean_data, csv_file_path)

        messagebox.showinfo("Process Completed", "XML processing is completed successfully!")


def save_clean_data(clean_data, file_path):
    # Create a DataFrame from the clean_data
    df = pd.DataFrame(clean_data)

    # Save the DataFrame to an Excel file
    df.to_excel(file_path, index=False)


def extract_time_values_from_xml(xml_file):
    with open(xml_file, 'r') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'xml')

    time_from = soup.find('TIME_FROM').text.strip()
    time_until = soup.find('TIME_UNTIL').text.strip()

    return time_from, time_until


def save_clean_data_to_csv(clean_data, file_path):
    # Open the file in write mode with newline='' to prevent extra line breaks
    with open(file_path, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(clean_data[0].keys())  # Write the header row with the column names
        writer.writerows([data.values() for data in clean_data])  # Write the data rows


def load_xml_file():
    file_path = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
    if file_path:
        column_index = column_index_var.get()
        process_xml_file(file_path, column_index)


# Cria a janela principal
window = tk.Tk()

# Variável para armazenar o índice da coluna selecionada
column_index_var = tk.IntVar()
column_index_var.set(1)

# Cria um botão para carregar o arquivo XML
load_button = tk.Button(window, text="Carregar arquivo XML", command=load_xml_file)
load_button.pack()

# Cria uma lista de opções para selecionar o índice da coluna
column_index_label = tk.Label(window, text="Selecione o índice da coluna:")
column_index_label.pack()

column_index_options = [(0, "Columns - 1 - kW"), (1, "Columns - 2 kvar"), (2, "Columns - 3 kvar"),
                        (3, "Columns - 4 - kWh"), (4, "Columns - 5 kvar"), (5, "Columns - 6 kvar")]

for option in column_index_options:
    rb = tk.Radiobutton(window, text=option[1], variable=column_index_var, value=option[0])
    rb.pack(anchor=tk.W)

# Inicia o loop da interface gráfica
window.mainloop()
