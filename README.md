# Riley's Portfolio
EVE Online Market Tool (Python)


Description:


A margin trading tool for the video game EVE Online's main trading hub with a customizeable UI and sort/filter features. Cross referencing the queried item name with an item ID CSV file as a database, it queries the market API information for the corresponding item ID and returns it in a customizeable UI, saving the queried information as a CSV with timestamp of data for API call optimization. Utilizes Pandas, EVE Online's web API, and tkinter for the heavy lifting.


Requirements:


Please download the following libraries if they are not already installed. This can be done by typing "pip install [library name]" into command prompt.

pandas

requests


How to Use:
Download the script and the required item_id.csv file. Extract them into the same directory. Run the program and type an item from EVE Online to search it and return market data. Stored market data resides in a created saved_market_data.csv file.


