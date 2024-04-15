import sys
import os
import cv2
import easyocr
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import urllib.request
import urllib.request
from bs4 import BeautifulSoup
import pandas as pd
from PIL import Image
from io import BytesIO



# Function to handle captcha processing
def process_captcha(browser):
    # Get the captcha image element
    captcha_image = browser.find_element(By.XPATH, '//*[@id="captcha"]')


    # Save the captcha image locally
    with open('captcha.png', 'wb') as file:
        file.write(captcha_image.screenshot_as_png)
        time.sleep(3)
    
    # Redirect stdout to a file for progress output
    # with open('output.log', 'w', encoding='utf-8') as f:
    #     sys.stdout = f
    #     reader = easyocr.Reader(['en'])
    #     sys.stdout = sys.__stdout__
   
    # Load the reader
    reader = easyocr.Reader(['en'])
    result = reader.readtext('captcha.png', detail=0)
    captcha_text = ''.join(result)
    print("captcha_text", captcha_text)

    return captcha_text



# Function to scrape data from webpage
def scrape_data(browser):
    try:
        # Parsing HTML content of the webpage
        soup = BeautifulSoup(browser.page_source, 'html.parser')

        # Scraping data from div elements with multiple rows
        div_elements = soup.select('div.card-body')
        div_details = []
        for div_element in div_elements:
            rows = div_element.select('div.row')
            details = {}
            for row in rows:
                label = row.select_one('div.font-12.font-weight-semi-bold')
                value = row.select_one('div.font-12.text-gray')
                if label and value:
                    details[label.text.strip()] = value.text.strip()
            div_details.append(details)

        # Scraping data from table format
        table = soup.select_one('table.table.table-hover.custom-datatable.dataTable.no-footer')
        table_details = []
        if table:
            rows = table.select('tr')
            header_row = rows[0]
            headers = [header.text.strip() for header in header_row.select('th')]
            for row in rows[1:]:
                cells = row.select('td')
                if cells:
                    row_data = [cell.text.strip() for cell in cells]
                    table_details.append(dict(zip(headers, row_data)))

        # Convert scraped data to DataFrame
        div_df = pd.DataFrame(div_details)
        print("Div Data:\n", div_df)
        table_df = pd.DataFrame(table_details)
        print("\nTable Data:\n", table_df)

        # Define the path to save the Excel file
        excel_file = "scraped_data.xlsx"

        # Create a Pandas Excel writer using XlsxWriter as the engine
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            # Write each DataFrame to a separate worksheet
            div_df.to_excel(writer, sheet_name='Div Data', index=False)
            table_df.to_excel(writer, sheet_name='Table Data', index=False)

        print(f"\nData has been exported to {excel_file}")
        return True

    except Exception as e:
        traceback.print_exc()
        print("Error occurred while scraping data:", e)
        return False


# Disable printing progress information during download
def no_progress(blocknum, bs, size):
    pass

# Redirect stdout to a file
sys.stdout = open(os.devnull, 'w')

chrome_options = webdriver.ChromeOptions()
browser = webdriver.Chrome(options=chrome_options)
browser.get('https://www.dgft.gov.in/CP/?opt=view-any-ice')

try:
    next_xpath = browser.find_element(By.XPATH, '//*[@id="mainSectionWrap"]/div[3]/div/div[2]/div[1]')
    next_xpath.click()
    time.sleep(2)

    iec_input = browser.find_element(By.XPATH, '//*[@id="iecNo"]')
    entity_input = browser.find_element(By.XPATH, '//*[@id="entity"]')
    iec_input.send_keys("0301014175")
    time.sleep(3)
    entity_input.send_keys("ACCENTURE SERVICES PVT. LTD.")
    time.sleep(3)
    
    
    soup = BeautifulSoup(browser.page_source, 'html.parser')
    # Print the HTML content to the terminal
    print(soup.prettify())
    
    while True:
        # Process the captcha
        captcha_text = process_captcha(browser)
        print("captcha_text", captcha_text)
        
        # Input the recognized captcha text
        captcha_input = browser.find_element(By.XPATH, '//*[@id="txt_Captcha"]')
        captcha_input.send_keys(captcha_text)

        time.sleep(3)

        # Click on the element with XPath '//*[@id="viewIEC1"]'
        view_button = browser.find_element(By.XPATH, '//*[@id="viewIEC1"]')
        view_button.click()
        time.sleep(5)

        # Check if the "Please enter valid captcha code" message is displayed
        error_message = browser.find_elements(By.XPATH, '//*[@id="incCaptcha"]')
        if not error_message:
            if scrape_data(browser):
                    break  # Exit loop if scraping successful



except Exception as e:
    traceback.print_exc()
    print("wrong xpath", e)

finally:
    browser.quit()

    # Restore stdout  
    sys.stdout = sys.__stdout__








# import sys
# import os
# import cv2
# import easyocr
# import traceback
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# import time
# import urllib.request
# from bs4 import BeautifulSoup
# import pandas as pd

# # Function to handle captcha processing
# def process_captcha(browser):
#     # Get the captcha image element
#     captcha_image = browser.find_element(By.XPATH, '//*[@id="captcha"]')

#     # Save the captcha image locally
#     with open('captcha.png', 'wb') as file:
#         file.write(captcha_image.screenshot_as_png)


#     # Redirect stdout to a file for progress output
#     with open('output.log', 'w', encoding='utf-8') as f:
#         sys.stdout = f
        
        
#     # Preprocess the captcha image
#     captcha_image = cv2.imread('captcha.png')
#     gray_image = cv2.cvtColor(captcha_image, cv2.COLOR_BGR2GRAY)
#     # _, binary_image = cv2.threshold(gray_image, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)
#     # blurred_image = cv2.GaussianBlur(binary_image, (3, 3), 0)
#     cv2.imwrite('preprocessed_captcha.png', gray_image)

#     # Load the reader
#     reader = easyocr.Reader(['en'])
#     result = reader.readtext('captcha.png', detail=0)
#     captcha_text = ''.join(result)
#     print("captcha_text", captcha_text)

#     return captcha_text

# # Disable printing progress information during download
# def no_progress(blocknum, bs, size):
#     pass

# # Redirect stdout to a file
# sys.stdout = open(os.devnull, 'w')

# chrome_options = webdriver.ChromeOptions()
# browser = webdriver.Chrome(options=chrome_options)
# browser.get('https://www.dgft.gov.in/CP/?opt=view-any-ice')

# try:
#     next_xpath = browser.find_element(By.XPATH, '//*[@id="mainSectionWrap"]/div[3]/div/div[2]/div[1]')
#     next_xpath.click()
#     time.sleep(2)

#     iec_input = browser.find_element(By.XPATH, '//*[@id="iecNo"]')
#     entity_input = browser.find_element(By.XPATH, '//*[@id="entity"]')
#     iec_input.send_keys("0301014175")
#     time.sleep(2)
#     entity_input.send_keys("ACCENTURE SERVICES PVT. LTD.")
#     time.sleep(2)

#     attempts = 0  # Initialize attempts counter
   
#     captcha_text = process_captcha(browser)
   
#     while True:
#         attempts += 1  # Increment attempts counter
#         print(f"Attempt {attempts} to enter captcha")  # Print attempt number

#         # Process the captcha
      

#         # Input the recognized captcha text
#         captcha_input = browser.find_element(By.XPATH, '//*[@id="txt_Captcha"]')
#         captcha_input.send_keys(captcha_text)

#         time.sleep(2)

#         # Click on the element with XPath '//*[@id="viewIEC1"]'
#         view_button = browser.find_element(By.XPATH, '//*[@id="viewIEC1"]')
#         view_button.click()
#         time.sleep(3)

#         # Check if the "Please enter valid captcha code" message is displayed
#         error_message = browser.find_elements(By.XPATH, '//*[@id="incCaptcha"]')
#         if not error_message:
#             print("Captcha entered successfully")  # Print success message
            
#             print("Parsing HTML content of the webpage...")
               
#             # Parse the HTML content of the webpage
#             soup = BeautifulSoup(browser.page_source, 'html.parser')
#             print("soup ", soup)

#             # Scraping data from div elements with multiple rows
#             div_elements = soup.find_all('div', class_='card-body')
#             div_details = []
#             for div_element in div_elements:
#                 rows = div_element.find_all('div', class_='row')
#                 details = {}
#                 for row in rows:
#                     label = row.find('div', class_='font-12 font-weight-semi-bold').text.strip()
#                     value = row.find('div', class_='font-12 text-gray ').text.strip()
#                     details[label] = value
#                 div_details.append(details)

#             # Scraping data from table format
#             table = soup.find('table', class_='table table-hover custom-datatable dataTable no-footer')
#             table_details = []
#             if table:
#                 rows = table.find_all('tr')
#                 for row in rows:
#                     cells = row.find_all('td')
#                     if cells:
#                         table_details.append([cell.text.strip() for cell in cells])

#             # Convert scraped data to DataFrame
#             div_df = pd.DataFrame(div_details)
#             print("div_df :" , div_df)
#             table_df = pd.DataFrame(table_details)
#             print("table_df :" , table_df)

#             # # Define the path to save the Excel file
#             excel_file = "scraped_data.xlsx"

#             # Create a Pandas Excel writer using XlsxWriter as the engine
#             with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
#                 # Write each DataFrame to a separate worksheet
#                 div_df.to_excel(writer, sheet_name='Div Data', index=False)
#                 table_df.to_excel(writer, sheet_name='Table Data', index=False)

#             print(f"Data has been exported to {excel_file}")
            
            
            
            
            
#             break  # Break the loop if no error message is found

  

# except Exception as e:
#     traceback.print_exc()
#     print("wrong xpath", e)

# finally:
#     browser.quit()

#     # Restore stdout
#     sys.stdout = sys.__stdout__



# import sys
# import os
# import cv2
# import easyocr
# import traceback
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# import time
# import urllib.request

# # Disable printing progress information during download
# def no_progress(blocknum, bs, size):
#     pass

# # Redirect stdout to a file
# sys.stdout = open(os.devnull, 'w')

# chrome_options = webdriver.ChromeOptions()
# browser = webdriver.Chrome(options=chrome_options)
# browser.get('https://www.dgft.gov.in/CP/?opt=view-any-ice')

# try:
#     next_xpath = browser.find_element(By.XPATH, '//*[@id="mainSectionWrap"]/div[3]/div/div[2]/div[1]')
#     next_xpath.click()
#     time.sleep(2)

#     iec_input = browser.find_element(By.XPATH, '//*[@id="iecNo"]')
#     entity_input = browser.find_element(By.XPATH, '//*[@id="entity"]')
#     iec_input.send_keys("0301014175")
#     time.sleep(3)
#     entity_input.send_keys("ACCENTURE SERVICES PVT. LTD.")
#     time.sleep(3)

#     # Get the captcha image element
#     captcha_image = browser.find_element(By.XPATH, '//*[@id="captcha"]')

#     # Save the captcha image locally
#     with open('captcha.png', 'wb') as file:
#         file.write(captcha_image.screenshot_as_png)

#     # Redirect stdout to a file for progress output
#     with open('output.log', 'w', encoding='utf-8') as f:
#         sys.stdout = f
        
#         # Preprocess the captcha image
#         captcha_image = cv2.imread('captcha.png')
#         gray_image = cv2.cvtColor(captcha_image, cv2.COLOR_BGR2GRAY)
#         _, binary_image = cv2.threshold(gray_image, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)
#         blurred_image = cv2.GaussianBlur(binary_image, (3, 3), 0)
#         cv2.imwrite('preprocessed_captcha.png', blurred_image)

#         # Perform OCR on the preprocessed image
#         reader = easyocr.Reader(['en'])
#         result = reader.readtext('preprocessed_captcha.png', detail=0)
#         captcha_text = ''.join(result)
#         print("captcha_text", captcha_text)

#     time.sleep(5)

#     # Input the recognized captcha text
#     captcha_input = browser.find_element(By.XPATH, '//*[@id="txt_Captcha"]')
#     captcha_input.send_keys(captcha_text)

#     time.sleep(3)

# except Exception as e:
#     traceback.print_exc()
#     print("wrong xpath", e)

# finally:
#     browser.quit()

#     # Restore stdout
#     sys.stdout = sys.__stdout__










# import sys
# import os
# import cv2
# import easyocr
# import traceback
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# import time
# import urllib.request

# # Disable printing progress information during download
# def no_progress(blocknum, bs, size):
#     pass

# # Redirect stdout to a file
# sys.stdout = open(os.devnull, 'w')

# chrome_options = webdriver.ChromeOptions()
# browser = webdriver.Chrome(options=chrome_options)
# browser.get('https://www.dgft.gov.in/CP/?opt=view-any-ice')

# try:
#     next_xpath = browser.find_element(By.XPATH, '//*[@id="mainSectionWrap"]/div[3]/div/div[2]/div[1]')
#     next_xpath.click()
#     time.sleep(2)

#     iec_input = browser.find_element(By.XPATH, '//*[@id="iecNo"]')
#     entity_input = browser.find_element(By.XPATH, '//*[@id="entity"]')
#     iec_input.send_keys("0301014175")
#     time.sleep(3)
#     entity_input.send_keys("ACCENTURE SERVICES PVT. LTD.")
#     time.sleep(3)

#     # Get the captcha image element
#     captcha_image = browser.find_element(By.XPATH, '//*[@id="captcha"]')

#     # Save the captcha image locally
#     with open('captcha.png', 'wb') as file:
#         file.write(captcha_image.screenshot_as_png)

#     # Redirect stdout to a file for progress output
#     with open('output.log', 'w', encoding='utf-8') as f:
#         sys.stdout = f
#         reader = easyocr.Reader(['en'])
#         sys.stdout = sys.__stdout__
   
#     # Load the reader
#     reader = easyocr.Reader(['en'])
#     result = reader.readtext('captcha.png', detail=0)
#     captcha_text = ''.join(result)
#     print("captcha_text", captcha_text)

#     time.sleep(5)

#     # Input the recognized captcha text
#     captcha_input = browser.find_element(By.XPATH, '//*[@id="txt_Captcha"]')
#     captcha_input.send_keys(captcha_text)

#     time.sleep(3)
    

#     # Click on the element with XPath '//*[@id="viewIEC1"]'
#     view_button = browser.find_element(By.XPATH, '//*[@id="viewIEC1"]')
#     view_button.click()
#     time.sleep(15)

# except Exception as e:
#     traceback.print_exc()
#     print("wrong xpath", e)

# finally:
#     browser.quit()

#     # Restore stdout
#     sys.stdout = sys.__stdout__







# Function to handle captcha processing
# def process_captcha(browser):
#     # Get the captcha image element
#     captcha_image = browser.find_element(By.XPATH, '//*[@id="captcha"]')

#     # Take a screenshot of the entire browser window
#     screenshot = browser.get_screenshot_as_png()

#     # Convert the screenshot to an image
#     img = Image.open(BytesIO(screenshot))

#     # Get the position and size of the captcha image element
#     location = captcha_image.location
#     size = captcha_image.size

#     # Crop the captcha image from the screenshot
#     captcha_img = img.crop((location['x'], location['y'], 
#                             location['x'] + size['width'], location['y'] + size['height']))

#     # Save the captcha image locally
#     captcha_img.save('captcha.png')

#     # Load the reader
#     reader = easyocr.Reader(['en'])
#     result = reader.readtext('captcha.png', detail=0)
#     captcha_text = ''.join(result)
#     print("captcha_text", captcha_text)

#     return captcha_text
