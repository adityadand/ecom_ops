import time
import openpyxl
import pyautogui

# Function to read ASIN, SKU, HSN code from Excel file
def read_excel_data(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        asin, sku, hsn_code = row[:3]  # Assuming ASIN, SKU, HSN code are in first three columns
        data.append((asin, sku, hsn_code))
    return data

# Function to list product on Amazon
def list_product_on_amazon(asin, sku, hsn_code):
    print("ASIN:", asin)
    print("SKU:", sku)
    print("HSN Code:", hsn_code)

    # Select and delete any existing ASIN
    pyautogui.click(540, 772)  # Clicking on ASIN input field
    pyautogui.hotkey("ctrl", "a")  # Select all text
    pyautogui.press("delete")  # Delete selected text

    # Enter new ASIN
    pyautogui.write(asin)
    pyautogui.press("enter")
    time.sleep(2)  # Wait for page to load

    # Select condition as new
    pyautogui.click(1738, 652)  # Clicking on condition dropdown
    pyautogui.press("down", presses=2)
    pyautogui.press("enter")
    time.sleep(1)

    # Click on "Sell This Product"
    pyautogui.click(1762, 715)
    time.sleep(7)  # Wait for page to load

    # Fill Seller SKU
    print("Filling SKU:", sku)
    sellersku_location = pyautogui.locateOnScreen("sellersku.png")
    if sellersku_location:
        pyautogui.click(pyautogui.center(sellersku_location))
        pyautogui.write(sku)
    else:
        print("Seller SKU image not found")

    time.sleep(2)  # Adding a delay to ensure SKU is filled properly

    # Auto-fill Quantity
    quantity_location = pyautogui.locateOnScreen("quantity.png")
    if quantity_location:
        pyautogui.click(pyautogui.center(quantity_location))
        pyautogui.write("0")  # Auto-fill quantity to 0
    else:
        print("Quantity image not found")

    time.sleep(2)  # Adding a delay to ensure quantity is filled properly

  # Select Country/Region Of Origin
    country_dropdown_location = pyautogui.locateOnScreen("country_dropdown.png")
    if country_dropdown_location:
        pyautogui.click(country_dropdown_location[0] + 400, country_dropdown_location[1] + 20)  # Clicking on the dropdown
        time.sleep(1)
        pyautogui.click(country_dropdown_location[0] + 350 , country_dropdown_location[1] - 380)  # Clicking on the desired country/region
    else:
        print("Country dropdown image not found")
# Adding a delay to ensure country is selected properly

    # Fill HSN code
    print("Filling HSN Code:", hsn_code)
    hsn_code_str = str(hsn_code)
    hsn_location = pyautogui.locateOnScreen("hsn.png")
    if hsn_location:
        pyautogui.click(pyautogui.center(hsn_location))
        pyautogui.write(hsn_code_str)
    else:
        print("HSN image not found")

    # Wait for user to check the entry and continue
    input("Please verify the entry and press Enter to continue...")


# Main function
def main():
    excel_file_path = "productdata.xlsx"  # Change this to your Excel file path
    products = read_excel_data(excel_file_path)

    for product in products:
        asin, sku, hsn_code = product
        list_product_on_amazon(asin, sku, hsn_code)
        user_input = input("Do you want to list another product? (yes/no): ")
        if user_input.lower() != "yes":
            break

if __name__ == "__main__":
    main()
