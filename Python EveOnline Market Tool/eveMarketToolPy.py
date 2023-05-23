# EVE Online Station Trading Prospecting & Profit Estimation Program by Riley Knight (2023)



import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import DoubleVar
from tkinter import simpledialog
from tkinter import StringVar
import csv
import os
import pandas as pd
import requests
from functools import lru_cache
from datetime import datetime


# Calculations based off of station trading in Jita 4-4
JITA_STATION_ID = 60003760


# Load the CSV file with item names and item IDs
def load_item_id_dict_from_csv(file_path):
    df = pd.read_csv(file_path, header=0, encoding='ISO-8859-1')
    item_name_col = df.columns[0].strip()
    item_id_col = df.columns[1].strip()
    return dict(zip(df[item_name_col], df[item_id_col]))


# Get the item ID from the item name
def get_item_id(name, item_id_dict):
    name_lower = name.lower()
    for item_name, item_id in item_id_dict.items():
        if item_name.lower() == name_lower:
            return item_id
    return None


# Get the highest buy order, lowest sell order, total buy volume, and total sell volume for an item using API calls 
@lru_cache(maxsize=1024)
def get_item_prices(item_id):
    url = f"https://esi.evetech.net/latest/markets/10000002/orders?datasource=tranquility&order_type=all&type_id={item_id}"
    response = requests.get(url)
    if response.status_code == 200:
        orders = response.json()
        buy_orders = [order for order in orders if order["is_buy_order"]
                      and order["location_id"] == JITA_STATION_ID]
        sell_orders = [order for order in orders if not order["is_buy_order"]
                       and order["location_id"] == JITA_STATION_ID]
        highest_buy = max([order["price"]
                          for order in buy_orders], default=None)
        lowest_sell = min([order["price"]
                          for order in sell_orders], default=None)
        total_buy_volume = sum([order["volume_remain"]
                               for order in buy_orders])
        total_sell_volume = sum([order["volume_remain"]
                                for order in sell_orders])
        return (highest_buy, lowest_sell, total_buy_volume, total_sell_volume)
    return None, None, None, None


# Create a class for the Excel sheet application
class ExcelSheetApp(tk.Tk):
    def __init__(self):
        super().__init__()

        #Set the title and define variables as empty
        self.title("EVE Online Market Data")
        self.df = None
        self.detached_items = []
        self.min_station_margin = None

        # Create a notebook (tabbed interface)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True)

        # Create a Market Lookup tab
        self.market_lookup_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.market_lookup_tab, text="Market Lookup")

        # Load the CSV file with item names and item IDs
        self.item_id_dict = load_item_id_dict_from_csv('item_ids.csv')

        # Check if the saved market data CSV file exists and load it into the DataFrame
        saved_data_file = "saved_market_data.csv"
        if os.path.isfile(saved_data_file):
            self.df = pd.read_csv(saved_data_file)
        else:
            self.df = pd.DataFrame(columns=['Item Name', 'Item ID', 'Minimum Sell Order', 'Maximum Buy Order',
                                   'Profit Potential', 'Station Margin', 'Total Buy Volume', 'Total Sell Volume', 'Market Data Time'])



        # Update input frame parent to market_lookup_tab
        self.input_frame = ttk.Frame(self.market_lookup_tab)
        self.input_frame.pack(pady=10)

        # Add input label and entry
        self.item_label = ttk.Label(
            self.input_frame, text="Enter the item name:")
        self.item_label.grid(row=0, column=0, padx=(10, 5))

        # Definitions for the dropdown menu
        self.dropdown_open = tk.BooleanVar(value=False)
        self.after_id = None

        # Add item entry with autocomplete dropdown menu
        self.item_var = tk.StringVar()
        self.item_entry = ttk.Combobox(
            self.input_frame, textvariable=self.item_var, width=30)
        self.item_entry.grid(row=0, column=1, padx=(5, 10))
        self.item_entry.bind('<KeyRelease>', self.update_suggestions)

        # Add search button
        self.search_button = ttk.Button(
            self.input_frame, text="Search", command=self.search_item)
        self.search_button.grid(row=0, column=2, padx=(10, 10))

        # Add filter label and entry
        self.filter_label = ttk.Label(self.input_frame, text="Filter results:")
        self.filter_label.grid(row=1, column=0, padx=(10, 5))
        self.filter_entry = ttk.Entry(self.input_frame, width=30)
        self.filter_entry.grid(row=1, column=1, padx=(5, 10))
        self.filter_entry.bind('<KeyRelease>', self.filter_table)

        # Initialize the sorting state dictionary
        self.sort_state = {col: 0 for col in range(len(self.df.columns) + 1)}

        # Update table frame parent to market_lookup_tab
        self.table_frame = ttk.Frame(self.market_lookup_tab)
        self.table_frame.pack(fill='both', expand=True)

        # Create a Treeview widget
        self.table = ttk.Treeview(
            self.table_frame, columns=self.df.columns.tolist(), show='headings')
        self.table.pack(fill='both', expand=True)
        self.table['columns'] = self.df.columns.tolist()

        # Create column headers and set properties for the Treeview widget
        for i, col in enumerate(self.df.columns.tolist()):
            self.table.heading(
                i, text=col, command=lambda c=i: self.sort_table(c))
            self.table.column(i, anchor='center', width=120,
                              minwidth=100, stretch=tk.NO)

        # Insert data into the Treeview widget
        self.update_table()

        # Create a menu bar
        self.menu_bar = tk.Menu(self)
        self.config(menu=self.menu_bar)

        # Create toolbar
        self.toolbar = ttk.Frame(self, relief=tk.RAISED)
        self.toolbar.pack(side=tk.TOP, fill=tk.X)

        # Add toolbar buttons
        self.open_button = ttk.Button(
            self.toolbar, text="Color Code", command=self.color_code)
        self.open_button.pack(side=tk.LEFT, padx=2, pady=2)

        self.save_button = ttk.Button(
            self.toolbar, text="Show Only Profitable", command=self.show_only_profitable)
        self.save_button.pack(side=tk.LEFT, padx=2, pady=2)

        self.clear_filter_button = ttk.Button(
            self.toolbar, text="Clear Formatting", command=self.clear_formatting)
        self.clear_filter_button.pack(side=tk.LEFT, padx=2, pady=2)

        # Define the columns for the table
        self.columns = ["Item Name", "Minimum Sell Order", "Maximum Buy Order",
                        "Profit Potential", "Station Margin", "Total Buy Volume", "Total Sell Volume"]

        self.update_table()

    # Load the data from a CSV file
    def load_data_from_csv(self):
        saved_data_file = "saved_market_data.csv"
        if os.path.isfile(saved_data_file):
            self.df = pd.read_csv(saved_data_file)

    # Save the data to a CSV file
    def save_data_to_csv(self):
        self.df.to_csv("saved_market_data.csv", index=False)

    # Get the background color for the station margin
    def get_bg_color(self, value):
        if value >= 0:
            green_intensity = min(255, int(255 - abs(value) * 255))
            return f"#{green_intensity:02x}FF{green_intensity:02x}"
        else:
            red_intensity = min(255, int(255 - abs(value) * 255))
            return f"#{red_intensity:02x}FF{red_intensity:02x}"

    # Search for an item and update the table
    def search_item(self):
        item_name = self.item_entry.get()
        if not item_name:
            return

        # Make the search case-insensitive and get the item ID from the item name dictionary
        item_id = get_item_id(item_name.lower(), self.item_id_dict)
        if item_id:
            # Get the market data for the item
            highest_buy, lowest_sell, total_buy_volume, total_sell_volume = get_item_prices(item_id)
            # If the market data was found, update the table
            if highest_buy is not None and lowest_sell is not None:
                current_time = datetime.now().strftime('%Y-%m-%d %H:%M')
                profit_potential = lowest_sell - highest_buy
                station_margin = profit_potential / highest_buy
                new_data = {
                    'Item Name': [item_name],
                    'Item ID': [item_id],
                    'Minimum Sell Order': [lowest_sell],
                    'Maximum Buy Order': [highest_buy],
                    'Profit Potential': [profit_potential],
                    'Station Margin': [station_margin],
                    'Total Buy Volume': [total_buy_volume],
                    'Total Sell Volume': [total_sell_volume],
                    'Market Data Time': [current_time],
                }
                # Create a new DataFrame with the new data
                new_df = pd.DataFrame(new_data)
                # If the item is already in the DataFrame, update the row. Otherwise, add the new row to the DataFrame
                if item_id in self.df['Item ID'].values:
                    self.df.loc[self.df['Item ID'] == item_id] = new_df.values
                else:
                    self.df = pd.concat([self.df, new_df], ignore_index=True)
                # Update the table and save the data to a CSV file
                self.update_table()
                self.save_data_to_csv()
            # If the item name or market data was not found, show an error message
            else:
                messagebox.showerror(
                    "Error", "Market data not found for this item.")
        else:
            messagebox.showerror("Error", f"Item not found: {item_name}")

    # Flag the dropdown menu as open
    def dropdown_opened(self, event=None):
        self.dropdown_open = True

    # Flag the dropdown menu as closed
    def dropdown_closed(self, event=None):
        self.dropdown_open = False

    # Update the suggestions in the dropdown menu based on the user's input and schedule a function to open the dropdown menu after a delay of 500 milliseconds
    def update_suggestions(self, event=None):
        input_text = self.item_entry.get()

        if not input_text:
            self.item_entry['values'] = []
            return

        # Get the suggestions from the item name dictionary
        suggestions = [name for name in self.item_id_dict.keys(
        ) if input_text.lower() in name.lower()]

        # Limit the number of suggestions to 10
        max_suggestions = 10
        if len(suggestions) > max_suggestions:
            suggestions = suggestions[:max_suggestions]

        # Update the dropdown menu
        self.item_entry['values'] = suggestions

        # Cancel the function to open the dropdown menu if it exists 
        if self.after_id is not None:
            self.after_cancel(self.after_id)

        # Schedule a function to open the dropdown menu after a delay of 500 milliseconds 
        self.after_id = self.after(500, self.open_dropdown)

    # Open the dropdown menu if it is not already open
    def open_dropdown(self):
        if self.item_entry['values']:
            self.item_entry.event_generate('<Down>')

    # Update the table 
    def update_table(self, df=None):
        if df is not None:
            # Don't update the original DataFrame
            pass
        else:
            df = self.df

        # Clear the table and insert the data
        self.clear_table()
        # Calculate metrics for the provided DataFrame
        df = self.calculate_metrics()
        # Format the data for the table
        filtered_data = self.filter_data(df)
        formatted_data = self.format_data(filtered_data)
        # Insert the formatted data into the table
        self.insert_data_to_table(formatted_data)

    # Clear the table
    def clear_table(self):
        for row in self.table.get_children():
            self.table.delete(row)

    # Filter the data based on the user's input and calculate metrics
    def filter_data(self, df):
        query = self.filter_entry.get().lower()
        filtered_data = df[df['Item Name'].str.lower().str.contains(query)]

        if self.min_station_margin is not None:
            filtered_data = filtered_data[filtered_data['Station Margin']
                                          >= self.min_station_margin]

        return self.calculate_metrics(filtered_data)

    # Calculate metrics for the DataFrame
    def calculate_metrics(self, df=None):
        if df is None:
            df = self.df

        #Utilize data from the API to calculate metrics for the DataFrame
        if not df.empty:
            df['Profit Potential'] = df['Minimum Sell Order'] - \
                df['Maximum Buy Order']
            df['Station Margin'] = df['Profit Potential'] / \
                df['Maximum Buy Order']

        return df

    # Format the data for the table
    def format_data(self, df):
        formatted_data = []
        # Iterate through the DataFrame and format each row
        for index, row in df.iterrows():
            station_margin = row['Station Margin']
            formatted_row = [
                row["Item Name"],
                row["Item ID"],
                f"{row['Minimum Sell Order']:,.2f} ISK",
                f"{row['Maximum Buy Order']:,.2f} ISK",
                f"{row['Profit Potential']:,.2f} ISK",
                f"{station_margin:.2%}",
                f"{row['Total Buy Volume']:,.0f}",
                f"{row['Total Sell Volume']:,.0f}",
                f"{row['Market Data Time']}"
            ]
            # Add the formatted row to the list of formatted data
            formatted_data.append(formatted_row)

        return formatted_data

    # Insert the data into the table
    def insert_data_to_table(self, formatted_data):
        for row in formatted_data:
            tree_item = self.table.insert('', 'end', values=row)

    # Color code the table based on the station margin
    def color_code(self):
        for item in self.table.get_children():
            values = self.table.item(item)['values']
            # Get the station margin as a float
            station_margin = float(values[5].strip('%')) / 100
            bg_color = self.get_bg_color(station_margin)
            self.table.item(item, tags=(bg_color,))
            self.table.tag_configure(bg_color, background=bg_color)

    # Show only profitable items based on the user's input
    def show_only_profitable(self):
        min_margin = simpledialog.askfloat(
            "Show Only Profitable", "Enter the minimum Station Margin (%):", minvalue=0.0, maxvalue=100.0)
        if min_margin is not None and self.df is not None:
            self.min_station_margin = min_margin / 100
            self.update_table()

    # Clear the formatting for the table
    def clear_formatting(self):
        # Clear background color
        for item in self.table.get_children():
            self.table.item(item, tags=(""))
            self.table.tag_configure("", background="")

        # Reset the minimum station margin and update the table
        self.min_station_margin = None
        self.update_table()

    # Sort the table based on the column header
    def sort_table(self, c):
        if self.sort_state[c] == 0:
            # Set all other column sort states to 0 (not sorted)
            for col in self.sort_state:
                self.sort_state[col] = 0

        # Get the DataFrame column name
        column_name = self.columns[c]

        # Sort the DataFrame based on the column name and update the table
        self.columns = ["Item Name", "Item ID", "Minimum Sell Order", "Maximum Buy Order",
                        "Profit Potential", "Station Margin", "Total Buy Volume", "Total Sell Volume"]
        # Calculate the station margin if the column name is "Station Margin"
        if column_name == 'Station Margin':
            temp_df = self.df.copy()
            temp_df['Station Margin'] = (
                temp_df['Profit Potential']) / temp_df['Maximum Buy Order']
            if self.sort_state[c] == 0:
                self.df = temp_df.sort_values(by=column_name, ascending=True).drop(
                    columns=['Station Margin'])
                self.sort_state[c] = 1
            elif self.sort_state[c] == 1:
                self.df = temp_df.sort_values(by=column_name, ascending=False).drop(
                    columns=['Station Margin'])
                self.sort_state[c] = 2
            else:
                self.df = temp_df.drop(columns=['Station Margin'])
                self.sort_state[c] = 0
        # Sort the DataFrame based on the column name if the column name is not "Station Margin"
        else:
            if self.sort_state[c] == 0:
                self.df = self.df.sort_values(by=column_name, ascending=True)
                self.sort_state[c] = 1
            elif self.sort_state[c] == 1:
                self.df = self.df.sort_values(by=column_name, ascending=False)
                self.sort_state[c] = 2
            else:
                self.df = self.df.sort_index()
                self.sort_state[c] = 0

        # Update the table and the column headers with the sort state and sort icons
        self.update_table()

        for idx, col_name in enumerate(self.columns):
            if self.sort_state[idx] == 0:
                sort_icon = ""
            elif self.sort_state[idx] == 1:
                sort_icon = " \u25B2"  # Up arrow
            else:
                sort_icon = " \u25BC"  # Down arrow

            self.table.heading(idx, text=col_name + sort_icon,
                               command=lambda c=idx: self.sort_table(c))

    # Filter the table based on the user's input
    def filter_table(self, event):
        self.update_table()


# Run the application
if __name__ == "__main__":
    app = ExcelSheetApp()
    app.mainloop()
