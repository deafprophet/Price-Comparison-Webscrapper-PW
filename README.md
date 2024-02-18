Product Offer Finder
The Product Offer Finder is a Python-based tool that combines a graphical user interface (GUI) with web scraping capabilities to assist users in finding and recording the best product offers available online.

Features:
Search Google Shopping: Allows users to input search terms and retrieves recommended products from Google Shopping, saving the results to an Excel file.
Extract Offers: Given a list of product names from an Excel file, it finds the best available offers on Ceneo.pl and updates the Excel file with the offer links and prices.
User Interface: Provides a simple GUI for users to either upload a product list or search for recommended products without manual code execution.
How to Use:
Search for Products: Use the GUI to input a search term for Google Shopping. The results will be saved to "Recommendations.xlsx" in the script's directory.
Upload Product List: Select an existing Excel file with product names. The script will update this file with the best offer links and prices it finds online.
Review Results: Access the "Recommendations.xlsx" or the updated product list Excel file to view the best offers and their details.
Setup:
Ensure you have the required Python libraries installed:

bash
Copy code
pip install tkinter pandas openpyxl selenium webdriver_manager
Running the Script:
Execute the script in a Python environment:

bash
Copy code
python product_offer_finder.py
The GUI will appear, providing you with options to either upload an Excel file or search for recommended products.

Dependencies:
Python 3
Tkinter (for GUI)
Pandas (for data handling)
Openpyxl (for Excel file operations)
Selenium (for web scraping)
Webdriver_manager (for managing the browser driver)
