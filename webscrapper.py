import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import simpledialog
import pandas as pd
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException

def search_recommended_products(window):
    # Ask for user input to search for products.
    search_phrase = simpledialog.askstring("Search", "Enter a product to search for:", parent=window)
    if not search_phrase:
        return  # End function if no input.

    # Set up Selenium WebDriver for Chrome.
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    try:
        # Navigate to Google Shopping with the search query.
        driver.get(f"https://www.google.com/search?tbm=shop&q={search_phrase.replace(' ', '+')}")

        # Ensure product results are loaded.
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.sh-pr__product-results-grid.sh-pr__product-results'))
        )
        
        # Collect containers of individual products.
        product_containers = driver.find_elements(By.CSS_SELECTOR, 'div.sh-pr__product-results-grid.sh-pr__product-results > div.sh-dgr__gr-auto.sh-dgr__grid-result')

        # Prepare a list to hold product details.
        products = []
        
        # Extract name and link for each product.
        for container in product_containers:
            product_name = container.find_element(By.CSS_SELECTOR, 'div.EI11Pd > h3.tAxDx').text
            product_link = container.find_element(By.CSS_SELECTOR, 'a.Lq5OHe.eaGTj').get_attribute('href')
            products.append({"Product Name": product_name, "Product Link": product_link})

        # Convert the list to a DataFrame and save as Excel.
        pd.DataFrame(products).to_excel("Recommendations.xlsx", index=False)

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        driver.quit()  # Clean up driver resources.

def upload_excel():
    # Trigger the main script function.
    main()

def main_window():
    # Set up the main GUI window.
    window = tk.Tk()
    window.title("Product Offer Finder")
    window.geometry("400x200+500+300")

    # Create and place buttons for uploading Excel files and searching for products.
    upload_button = tk.Button(window, text="Upload Excel File with Products", command=lambda: [window.destroy(), upload_excel()])
    upload_button.pack(pady=10)
    search_button = tk.Button(window, text="Search for a List of Recommended Products", command=lambda: search_recommended_products(window))
    search_button.pack(pady=10)

    window.mainloop()  # Start the GUI event loop.

def select_excel_file():
    # Prompt the user to choose an Excel file.
    root = tk.Tk()
    root.withdraw()  # Temporarily hide the main window.
    file_path = askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx"), ("Excel files", "*.xls")])
    root.destroy()  # Close the temporary window.
    if not file_path:
        print("No file selected. Exiting script.")
        exit()  # Exit if no file is chosen.
    return file_path


def load_product_names(file_path):
    """
    Extracts a list of product names from the first column of an Excel file.
    """
    df = pd.read_excel(file_path)
    return df.iloc[:, 0].tolist()

def find_best_offer(driver, product_name):
    """Uses Selenium to find the best offer for a given product name on Ceneo.pl."""
    search_url = 'https://www.ceneo.pl/;szukaj-' + product_name.replace(" ", "+")
    driver.get(search_url)

    try:
        # Wait for the search results to load
        WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.CSS_SELECTOR,
                                            'div.category-list-body.js_category-list-body.js_search-results.js_products-list-main.js_async-container'))
        )

        # Locate the first product and its 'btn-compare-outer' container
        first_product_container = driver.find_element(By.CSS_SELECTOR,
                                                      'div.category-list-body.js_category-list-body.js_search-results.js_products-list-main.js_async-container > :first-child')
        btn_compare_outer = first_product_container.find_element(By.CSS_SELECTOR, '.btn-compare-outer')
        anchor = btn_compare_outer.find_element(By.TAG_NAME, 'a')

        # Extract the href attribute and the anchor text
        offer_link = anchor.get_attribute('href')
        anchor_text = anchor.text.strip().upper()  # Convert text to uppercase and strip whitespace

        # Conditional logic based on the anchor text
        if anchor_text == "IDŹ DO SKLEPU":
            # Find the price container and extract the price parts
            price_container = first_product_container.find_element(By.CSS_SELECTOR, 'div.cat-prod-row__price')
            price_value = price_container.find_element(By.CSS_SELECTOR, '.value').text
            price_penny = price_container.find_element(By.CSS_SELECTOR, '.penny').text
            full_price = f"{price_value}{price_penny}"
            return offer_link, full_price
        elif anchor_text == "PORÓWNAJ CENY":
            # Navigate to the modified link
            comparison_page_link = offer_link + ";0280-0.htm"
            driver.get(comparison_page_link)

            # Wait for the new page to load and for the product offers to be present
            WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.CSS_SELECTOR,
                                        'li.product-offers__list__item.js_productOfferGroupItem'))
            )

            # Find the first product offer link and price container
            first_offer_container = driver.find_element(By.CSS_SELECTOR,
                                                'li.product-offers__list__item.js_productOfferGroupItem')
            product_offer_link = first_offer_container.find_element(By.CSS_SELECTOR,
                                                           'div.product-offer__container.js_product-offer').find_element(By.TAG_NAME, 'a').get_attribute('href')

            # Find the price container within the first offer
            price_container = first_offer_container.find_element(By.CSS_SELECTOR, 'span.price-format.nowrap')
            price_value = price_container.find_element(By.CSS_SELECTOR, 'span.value').text
            price_penny = price_container.find_element(By.CSS_SELECTOR, 'span.penny').text
            full_price = f"{price_value}{price_penny}"

            return product_offer_link, full_price  # Return the product offer link and the scraped price
        else:
            print(f"Unexpected anchor text '{anchor_text}' for {product_name}")
            return None, None  # Return None for both offer link and price

    except NoSuchElementException:
        # If the first offer or necessary elements are not found, return indicators for "No offers found"
        print(f"No offers found for {product_name}")
        return "No offers found", "N/A"  # Indicate absence of offers in the Excel file

    except TimeoutException:
        print(f"Timeout waiting for offers for {product_name}")
        return "No offers found", "N/A"  # Handle timeout by indicating absence of offers

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return "No offers found", "N/A"  # Handle any other exceptions similarly

def update_excel_file(file_path, product_offers):
    """
    Updates the Excel file with the links and prices of product offers.

    :param file_path: Path to the Excel file that needs to be updated.
    :param product_offers: A list of tuples containing the product name, offer link, and price.
    """
    # Load the workbook and select the active worksheet to update.
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    
    # Iterate over the product offers, starting from row 2 (assuming row 1 has headers).
    for row, (product, link, price) in enumerate(product_offers, start=2):
        # If a link is present, write it to column B of the current row.
        if link: 
            sheet[f'B{row}'] = link
        # If a price is present, write it to column C of the current row.
        if price: 
            sheet[f'C{row}'] = price
    
    # After all product offers have been written, save the workbook.
    workbook.save(file_path)

def update_excel_file(file_path, product_offers):
    """
    Inserts product offer links and prices into an Excel file.
    """
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    
    for row, (product, link, price) in enumerate(product_offers, start=2):
        # Write product link and price to their respective columns.
        sheet[f'B{row}'] = link or 'N/A'  # Use 'N/A' if link is None.
        sheet[f'C{row}'] = price or 'N/A'  # Use 'N/A' if price is None.
    
    workbook.save(file_path)  # Save the updates to the Excel file.

def main():
    """
    Orchestrates the workflow of the script.
    """
    excel_file_path = select_excel_file()  # Get the path to the Excel file.
    if not excel_file_path:
        print("No Excel file selected. Exiting script.")
        return

    product_names = load_product_names(excel_file_path)  # Extract product names from the Excel file.
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))  # Set up the WebDriver.
    product_offers = []  # Container for product offers.
    
    for product in product_names:
        try:
            offer_details = find_best_offer(driver, product)  # Search for the best offer.
            product_offers.append((product,) + offer_details)  # Append product details to the list.
        except Exception as e:
            print(f"Error for {product}: {e}")
            product_offers.append((product, None, None))  # Append with None values on error.
    
    driver.quit()  # Terminate the WebDriver.
    update_excel_file(excel_file_path, product_offers)  # Update the Excel file with offer details.
    print("Script completed. Excel file updated.")

# Python boilerplate for script execution.
if __name__ == "__main__":
    main_window()
