from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import os
import time

# Initialize libraries
browser = Selenium()
http = HTTP()
tables = Tables()
pdf = PDF()
archive = Archive()

# Configuration parameters
URL = "https://robotsparebinindustries.com/#/robot-order"
ORDERS_URL = "https://robotsparebinindustries.com/orders.csv"
ORDERS_FILE = "orders.csv"
OUTPUT_DIR = "output"
RETRY_LIMIT = 3

# Ensure output directories exist
os.makedirs(f"{OUTPUT_DIR}/screenshot", exist_ok=True)
os.makedirs(f"{OUTPUT_DIR}/pdf", exist_ok=True)

# Mapping for robot parts
head_mapping = {
    "1": "Roll-a-thor head",
    "2": "Peanut crusher head",
    "3": "D.A.V.E head",
    "4": "Andy Roid head",
    "5": "Spanner mate head",
    "6": "Drillbit 2000 head"
}

body_mapping = {
    "1": "Roll-a-thor body",
    "2": "Peanut crusher body",
    "3": "D.A.V.E body",
    "4": "Andy Roid body",
    "5": "Spanner mate body",
    "6": "Drillbit 2000 body"
}

def open_the_intranet_website():
    """Navigates to the given URL"""
    print("Opening the intranet website...")
    browser.open_available_browser(URL)
    print("Website opened.")

def get_orders():
    """Downloads the orders CSV file and reads it into a table."""
    print("Downloading orders CSV file...")
    http.download(url=ORDERS_URL, target_file=ORDERS_FILE, overwrite=True)
    orders_table = tables.read_table_from_csv(ORDERS_FILE)
    print("Orders CSV file downloaded and read.")
    return orders_table

def close_annoying_modal():
    """Closes the modal if it appears"""
    try:
        ok_button_locator = "xpath://*[@id='root']/div/div[2]/div/div/div/div/div/button[1]"
        if browser.is_element_visible(ok_button_locator):
            browser.click_element(ok_button_locator)
            browser.wait_until_element_is_not_visible(ok_button_locator, timeout=10)
            print("Modal closed successfully.")
        else:
            print("Modal not found.")
    except Exception as e:
        print(f"Error closing modal: {e}")

def fill_the_form(order_data):
    """Fills the form with the given order details"""
    try:
        head_model = head_mapping[order_data['Head']]
        browser.select_from_list_by_label("//select[@id='head']", head_model)
        
        body_model_label = body_mapping[order_data['Body']]
        body_radio_xpath = f"//label[contains(text(), '{body_model_label}')]/input"
        browser.click_element(body_radio_xpath)
        
        browser.input_text("//label[contains(text(), 'Legs:')]/following-sibling::input", order_data['Legs'])
        browser.input_text("//input[@id='address']", order_data['Address'])
        print("Form filled successfully.")
    except Exception as e:
        print(f"An error occurred while filling the form: {e}")

def preview_the_robot():
    """Previews the robot"""
    try:
        browser.click_button("preview")
        browser.wait_until_element_is_visible("xpath://*[@id='robot-preview-image']", timeout=10)
        print("Robot previewed successfully.")
    except Exception as e:
        print(f"An error occurred while previewing the robot: {e}")

def submit_the_order():
    """Submits the order"""
    try:
        browser.click_button("order")
        browser.wait_until_element_is_visible("xpath://*[@id='receipt']", timeout=10)
        print("Order submitted successfully.")
    except Exception as e:
        print(f"An error occurred while submitting the order: {e}")

def take_screenshot_of_robot(order_number):
    """Takes a screenshot of the robot preview"""
    try:
        robot_preview_xpath = "//*[@id='robot-preview-image']"
        screenshot_path = f'{OUTPUT_DIR}/screenshot/robot_preview_image_{order_number}.png'
        browser.capture_element_screenshot(robot_preview_xpath, screenshot_path)
        if os.path.exists(screenshot_path):
            print(f"Screenshot taken successfully for order {order_number}.")
        else:
            print(f"Screenshot was not saved for order {order_number}.")
        return screenshot_path
    except Exception as e:
        print(f"An error occurred while taking the screenshot: {e}")
        return None

def store_receipt_as_pdf(order_number):
    """Stores the receipt as a PDF file"""
    try:
        receipt_xpath = "//*[@id='receipt']"
        receipt_element = browser.find_element(receipt_xpath)
        receipt_html = receipt_element.get_attribute("outerHTML")
        receipt_filename = f'{OUTPUT_DIR}/pdf/receipt_{order_number}.pdf'
        pdf.html_to_pdf(receipt_html, receipt_filename)
        if os.path.exists(receipt_filename):
            print(f"Receipt PDF saved successfully for order {order_number}.")
        else:
            print(f"Receipt PDF was not saved for order {order_number}.")
        return receipt_filename
    except Exception as e:
        print(f"An error occurred while saving the receipt for order {order_number}: {e}")
        return None

def embed_robot_screenshot_to_pdf(screenshot, pdf_file):
    """Embeds the robot screenshot into the PDF"""
    try:
        pdf.add_watermark_image_to_pdf(screenshot, pdf_file, pdf_file)
        print("Screenshot embedded into PDF successfully.")
    except Exception as e:
        print(f"An error occurred while embedding the screenshot to the PDF: {e}")

def go_to_order_another_robot():
    """Navigates to order another robot"""
    try:
        order_another_button_xpath = "xpath://*[@id='order-another']"
        browser.click_element(order_another_button_xpath)
        print("Navigated to order another robot.")
    except Exception as e:
        print(f"An error occurred while trying to order another robot: {e}")

def create_zip_file_of_receipts():
    """Creates a ZIP file of all receipts"""
    try:
        zip_file_name = f'{OUTPUT_DIR}/all_receipts.zip'
        archive.archive_folder_with_zip(f'{OUTPUT_DIR}/pdf', zip_file_name)
        print(f"Receipts archived to {zip_file_name}")
    except Exception as e:
        print(f"An error occurred while creating the ZIP file: {e}")

def process_order(row):
    """Processes a single order with retry logic for each step"""
    order_number = row['Order number']
    print(f"Processing order {order_number}...")

    success = False
    for attempt in range(RETRY_LIMIT):
        try:
            close_annoying_modal()
            fill_the_form(row)
            preview_the_robot()
            submit_the_order()
            pdf_file = store_receipt_as_pdf(order_number)
            if pdf_file:
                screenshot_file = take_screenshot_of_robot(order_number)
                if screenshot_file:
                    embed_robot_screenshot_to_pdf(screenshot_file, pdf_file)
            go_to_order_another_robot()
            success = True
            break
        except Exception as e:
            print(f"An error occurred for order {order_number} on attempt {attempt + 1}: {e}")
            time.sleep(2)  # Optional: Delay before retrying

    if not success:
        print(f"Failed to process order {order_number} after {RETRY_LIMIT} attempts")

@task
def order_robots_from_RobotSpareBin():
    """Main task to order robots from Robot Spare Bin"""
    try:
        open_the_intranet_website()
        orders = get_orders()
        for row in orders:
            process_order(row)
        create_zip_file_of_receipts()
    except Exception as e:
        print(f"An error occurred in the order_robots_from_RobotSpareBin task: {e}")
    finally:
        browser.close_all_browsers()
