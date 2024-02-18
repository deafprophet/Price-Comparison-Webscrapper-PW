# Web Scrapper

The Web Scrapper is a Python-based tool that leverages a graphical user interface (GUI) with web scraping capabilities to assist users in finding and documenting the best product offers online.

## Features

- **Search Google Shopping**: Users can enter search terms to retrieve recommended products from Google Shopping, which are then saved to an Excel file.
- **Extract Offers**: Processes a list of product names from an Excel file to find the best offers on Ceneo.pl, updating the file with the details.
- **User Interface**: Offers a simple GUI for users to either upload a product list or initiate a product search, facilitating ease of use without manual code intervention.

## How to Use

1. **Search for Products**: Through the GUI, enter a search term for Google Shopping. Results are saved to "Recommendations.xlsx" in the script's directory.
2. **Upload Product List**: Use the GUI to select an Excel file with product names. The script will find the best offers online and update this file accordingly.
3. **Review Results**: Check "Recommendations.xlsx" or the updated Excel file for detailed offer information.

## Setup

Install the required Python libraries:

```python
pip install tkinter pandas openpyxl selenium webdriver_manager
```

## Running the Script
To run the script, use the following command in a Python environment:

```python
python webscrapper.py
```

## Dependencies
- Python 3
- Tkinter (for the GUI)
- Pandas (for data manipulation)
- Openpyxl (for Excel operations)
- Selenium (for web scraping tasks)
- Webdriver_manager (for browser driver management)
