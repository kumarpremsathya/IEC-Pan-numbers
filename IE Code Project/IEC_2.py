
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
import sys
import os
import cv2
import easyocr
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import urllib.request
from selenium.common.exceptions import WebDriverException




import json
import sys
import os
import cv2
import easyocr
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from bs4 import BeautifulSoup
import pandas as pd



# Function to handle captcha processing
def solve_captcha(browser):
    captcha_text = process_captcha(browser)
    print("captcha_text " , captcha_text)
    
    captcha_input = browser.find_element(By.XPATH, '//*[@id="txt_Captcha"]')
    captcha_input.clear()
    captcha_input.send_keys(captcha_text)

    time.sleep(3)

    view_button = browser.find_element(By.XPATH, '//*[@id="viewIEC1"]')
    view_button.click()
    time.sleep(3)

     # Check if the "Please enter valid captcha code" message is displayed
    error_message = browser.find_element(By.XPATH, '//*[@id="incCaptcha"]').text
    print(" error_message",error_message) 
    
    if "Please enter valid captcha code" in error_message:
        return solve_captcha(browser)  # Retry solving captcha recursively
        
    else:
        return True


#Function to handle captcha processing
def process_captcha(browser):
    # Get the captcha image element
    captcha_image = browser.find_element(By.XPATH, '//*[@id="captcha"]')


    # Save the captcha image locally
    with open('captcha.png', 'wb') as file:
        file.write(captcha_image.screenshot_as_png)
        
    # Preprocess the captcha image
    captcha_image = cv2.imread('captcha.png')
    gray_image = cv2.cvtColor(captcha_image, cv2.COLOR_BGR2GRAY)
    # _, binary_image = cv2.threshold(gray_image, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)
    # blurred_image = cv2.GaussianBlur(binary_image, (3, 3), 0)
    cv2.imwrite('preprocessed_captcha.png', gray_image)
    
    # Load the reader
    reader = easyocr.Reader(['en'])
    result = reader.readtext('captcha.png', detail=0)
    captcha_text = ''.join(result)
    print("captcha_text", captcha_text)

    return captcha_text


def scrape_data(browser):
    try:
        
       
        symbol_click =  browser.find_element(By.XPATH, '//*[@id="custom-accordion"]/div[2]/div[1]/a')
        
        symbol_click.click
        
         # Locate the element containing the data using its XPath
        rcmc_element = browser.find_element(By.XPATH, '//*[@id="rcmc"]/div')
        print("rcmc_element :" , rcmc_element)
        
         # Get the HTML content of the element
        rcmc_html = rcmc_element.get_attribute('innerHTML')
        
        # Parse the HTML content using BeautifulSoup
        soup = BeautifulSoup(rcmc_html, 'html.parser')
        
        
        rcmc_branch_details = [] 
        
        # Scraping data from table format
        table = soup.find('table', class_='table table-hover custom-datatable')
        print("table :" , table)
        if table:
            headers = [th.text.strip() for th in table.find_all('th')]
            rows = table.find_all('tr')
            for row in rows:
                cells = row.find_all('td')
                if cells:
                    # Extract text from cells and add to all_branch_details
                    row_details = [cell.text.strip() for cell in cells]
                    rcmc_branch_details.append(dict(zip(headers, row_details)))
        
        # Convert all_branch_details to JSON string
        rcmc_details_json = json.dumps(rcmc_branch_details)

        # Create a Pandas DataFrame with the JSON data
        rcmc_df = pd.DataFrame({"RCMC DETAILS": [rcmc_details_json]})
        print("rcmc_df", rcmc_df)
        
        
        # Locate the element containing the data using its XPath
        data_element = browser.find_element(By.XPATH, '//*[@id="iecdetails"]')
        
        # Get the HTML content of the element
        data_html = data_element.get_attribute('innerHTML')
        
        # Parse the HTML content using BeautifulSoup
        soup = BeautifulSoup(data_html, 'html.parser')

        # Scraping data from div elements with multiple rows
        div_elements = soup.find_all('div', class_='card-body')
                      
        div_details = []
        for div_element in div_elements:
            rows = div_element.find_all('div', class_='form-group')
            details = {}
            for row in rows:
                label_element = row.find('label', class_='font-12 font-weight-semi-bold')
                if label_element:
                    label = label_element.text.strip()
                    value_element = row.find('p', class_='font-12 text-gray')
                    if value_element:
                        value = value_element.text.strip()
                        details[label] = value
            div_details.append(details)

        all_branch_details = []  # Initialize a list to store all branch details
        
        # Loop through each pagination link until there is no "Next" button
        while True:
            # Locate the element containing the data using its XPath
            data_element = browser.find_element(By.XPATH, '//*[@id="iecdetails"]')
            
            # Get the HTML content of the element
            data_html = data_element.get_attribute('innerHTML')
            
            # Parse the HTML content using BeautifulSoup
            soup = BeautifulSoup(data_html, 'html.parser')

            # Scraping data from table format
            table = soup.find('table', class_='table table-hover custom-datatable dataTable no-footer')
            if table:
                headers = [th.text.strip() for th in table.find_all('th')]
                rows = table.find_all('tr')
                for row in rows:
                    cells = row.find_all('td')
                    if cells:
                        # Extract text from cells and add to all_branch_details
                        row_details = [cell.text.strip() for cell in cells]
                        all_branch_details.append(dict(zip(headers, row_details)))
            
            # Click the "Next" button to navigate to the next page
            next_button = browser.find_element(By.ID, 'branchTable_next')
            if 'disabled' in next_button.get_attribute('class'):
                # If the "Next" button is disabled, break the loop
                break
            else:
                next_button.click()
                
                # Wait for the table to load (you may need to implement this)
                # Add code to wait for the table to load here
                time.sleep(2)  # Adjust the delay according to your page loading time
        
        # Convert all_branch_details to JSON string
        branch_details_json = json.dumps(all_branch_details)

        # Define the path to save the Excel file
        excel_file = "scraped_data.xlsx"

        # Create a Pandas DataFrame with the JSON data
        branch_df = pd.DataFrame({"BRANCH DETAILS": [branch_details_json]})
        
        # Write the DataFrame to an Excel file
        branch_df.to_excel(excel_file, index=False)
        
        print(f"Data has been exported to {excel_file}")
        
        # Convert scraped data to DataFrames
        table_df = pd.DataFrame(all_branch_details)
        div_df = pd.DataFrame(div_details)
        
        # Remove \n\t from DataFrame columns
        table_df = table_df.replace(r'\n\t','', regex=True)
        div_df = div_df.replace(r'\n\t','', regex=True)
        branch_df = branch_df.replace(r'\n\t','', regex=True)
        rcmc_df = rcmc_df.replace(r'\n\t','', regex=True)

        # Merge all DataFrames and write to Excel
        merged_df = pd.concat([pd.DataFrame({"IEC NUMBER": ["0301014175"]}), div_df, branch_df , rcmc_df], axis=1)
        with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
            merged_df.to_excel(writer, sheet_name='Merged Data', index=False)

        print(f"Data has been exported to {excel_file}")
            
        return True

    except Exception as e:
        print("Error occurred while scraping data:", e)
        return False


 # Function to scrape data from webpage
# def scrape_data(browser):
#     try:
#        # Locate the element containing the data using its XPath
#         data_element = browser.find_element(By.XPATH, '//*[@id="iecdetails"]')
        
#         # Get the HTML content of the element
#         data_html = data_element.get_attribute('innerHTML')
        
#         # Parse the HTML content using BeautifulSoup
#         soup = BeautifulSoup(data_html, 'html.parser')
#         # print(soup.prettify())
        

    
#         # Scraping data from div elements with multiple rows
#         div_elements = soup.find_all('div', class_='card-body')
                      
#         div_details = []
#         for div_element in div_elements:
#             rows = div_element.find_all('div', class_='form-group')
#             details = {}
#             for row in rows:
#                 label_element = row.find('label', class_='font-12 font-weight-semi-bold')
#                 if label_element:
#                     label = label_element.text.strip()
#                     value_element = row.find('p', class_='font-12 text-gray')
#                     if value_element:
#                         value = value_element.text.strip()
#                         details[label] = value
#             div_details.append(details)

#         all_branch_details = []  # Initialize a list to store all branch details
        
#         # Loop through each pagination link until there is no "Next" button
#         while True:
#             # Locate the element containing the data using its XPath
#             data_element = browser.find_element(By.XPATH, '//*[@id="iecdetails"]')
            
#             # Get the HTML content of the element
#             data_html = data_element.get_attribute('innerHTML')
            
#             # Parse the HTML content using BeautifulSoup
#             soup = BeautifulSoup(data_html, 'html.parser')

            
#             # Scraping data from table format
#             table = soup.find('table', class_='table table-hover custom-datatable dataTable no-footer')
#             if table:
#                 headers = [th.text.strip() for th in table.find_all('th')]
#                 rows = table.find_all('tr')
#                 for row in rows:
#                     cells = row.find_all('td')
#                     if cells:
#                         # Extract text from cells and add to all_branch_details
#                         row_details = [cell.text.strip() for cell in cells]
#                         all_branch_details.append(dict(zip(headers, row_details)))
            
#             # Click the "Next" button to navigate to the next page
#             next_button = browser.find_element(By.ID, 'branchTable_next')
#             if 'disabled' in next_button.get_attribute('class'):
#                 # If the "Next" button is disabled, break the loop
#                 break
#             else:
#                 next_button.click()
                
#                 # Wait for the table to load (you may need to implement this)
#                 # Add code to wait for the table to load here
#                 time.sleep(2)  # Adjust the delay according to your page loading time
        
#         # Convert all_branch_details to JSON string
#         branch_details_json = json.dumps(all_branch_details)

#         # Define the path to save the Excel file
#         excel_file = "scraped_data.xlsx"

#         # Create a Pandas DataFrame with the JSON data
#         df = pd.DataFrame({"BRANCH DETAILS": [branch_details_json]})
        
#         # Write the DataFrame to an Excel file
#         df.to_excel(excel_file, index=False)
        
#         print(f"Data has been exported to {excel_file}")
        
        
        
#         pd.set_option('display.max_columns', None)

#         # Convert scraped data to DataFrame
#         table_df = pd.DataFrame( all_branch_details)
       
#         print("table_df :" , table_df)
#         div_df = pd.DataFrame(div_details)
#         #Set display options to show all columns
#         pd.set_option('display.max_columns', None)
#         print("div_df :" , div_df)
       

#         # # Define the path to save the Excel file
#         # excel_file = "scraped_data.xlsx"

#         # # Create a Pandas Excel writer using XlsxWriter as the engine
#         # with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
#         #     #Write each DataFrame to a separate worksheet
#         #     div_df.to_excel(writer, sheet_name='Div Data', index=False)
#         #     table_df.to_excel(writer, sheet_name='Table Data', index=False)

#         print(f"Data has been exported to {excel_file}")
            
#         return True

#     except Exception as e:
#         traceback.print_exc()
#         print("Error occurred while scraping data:", e)
#         return False


# Disable printing progress information during download
def no_progress(blocknum, bs, size):
    pass

# Redirect stdout to a file
# sys.stdout = open(os.devnull, 'w')

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

    
    # Attempt to solve the captcha
    try:
        captcha_solved = solve_captcha(browser)
        print("captcha_solved:", captcha_solved)    
     
    except Exception as e:
          traceback.print_exc()
          print("captcha_solved:", e)    
     
    time.sleep(5)
     
    if captcha_solved:
        # soup = BeautifulSoup(browser.page_source, 'html.parser')
        # print("soup", soup)
        # print(soup.prettify())
        time.sleep(3)
        scrape_data(browser)
    else:
        print("Failed to solve captcha.")

except Exception as e:
    traceback.print_exc()
    print("An error occurred:", e)

finally:
    browser.quit()












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

#     # Loop until the correct captcha is entered
#     while True:
#         try:
#             # Get the captcha image element
#             captcha_image = browser.find_element(By.XPATH, '//*[@id="captcha"]')

#             # Save the captcha image locally
#             with open('captcha.png', 'wb') as file:
#                 file.write(captcha_image.screenshot_as_png)

#             # Redirect stdout to a file for progress output
#             with open('output.log', 'w', encoding='utf-8') as f:
#                 sys.stdout = f
#                 reader = easyocr.Reader(['en'])
#                 sys.stdout = sys.__stdout__
           
#             # Load the reader
#             reader = easyocr.Reader(['en'])
#             result = reader.readtext('captcha.png', detail=0)
#             captcha_text = ''.join(result)
#             print("captcha_text", captcha_text)

#             time.sleep(5)

#             # Input the recognized captcha text
#             captcha_input = browser.find_element(By.XPATH, '//*[@id="txt_Captcha"]')
#             captcha_input.send_keys(captcha_text)

#             time.sleep(3)

#             # Click on the element with XPath '//*[@id="viewIEC1"]'
#             view_button = browser.find_element(By.XPATH, '//*[@id="viewIEC1"]')
#             view_button.click()
#             time.sleep(3)

#             # Check if the "Please enter valid captcha code" message is displayed
#             error_message = browser.find_elements(By.XPATH, '//*[@id="incCaptcha"]')
#             if not error_message:
#                 break  # Break the loop if captcha entry was successful
        
#         except WebDriverException as e:
#             print("Error taking screenshot of captcha:", e)
#             # Retry taking screenshot or handle the error as necessary

# except Exception as e:
#     traceback.print_exc()
#     print("wrong xpath", e)

# finally:
#     browser.quit()

#     # Restore stdout
#     sys.stdout = sys.__stdout__





# # Function to handle captcha processing
# def process_captcha(browser):
#     # Get the captcha image element
#     captcha_image = browser.find_element(By.XPATH, '//*[@id="captcha"]')


#     # Save the captcha image locally
#     with open('captcha.png', 'wb') as file:
#         file.write(captcha_image.screenshot_as_png)
        
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



# # Function to scrape data from webpage
# def scrape_data(browser):
#     try:
#         # Parsing HTML content of the webpage
#         soup = BeautifulSoup(browser.page_source, 'html.parser')

        
#        #Scraping data from div elements with multiple rows
#         div_elements = soup.find_all('div', class_='card-body')
#         div_details = []
#         for div_element in div_elements:
#             rows = div_element.find_all('div', class_='row')
#             details = {}
#             for row in rows:
#                 label = row.find('div', class_='font-12 font-weight-semi-bold').text.strip()
#                 value = row.find('div', class_='font-12 text-gray ').text.strip()
#                 details[label] = value
#             div_details.append(details)

#         # Scraping data from table format
#         table = soup.find('table', class_='table table-hover custom-datatable dataTable no-footer')
#         table_details = []
#         if table:
#             rows = table.find_all('tr')
#             for row in rows:
#                 cells = row.find_all('td')
#                 if cells:
#                     table_details.append([cell.text.strip() for cell in cells])

#         # Convert scraped data to DataFrame
#         # div_df = pd.DataFrame(div_details)
#         # print("div_df :" , div_df)
#         table_df = pd.DataFrame(table_details)
#         print("table_df :" , table_df)

#         # # Define the path to save the Excel file
#         excel_file = "scraped_data.xlsx"

#         # Create a Pandas Excel writer using XlsxWriter as the engine
#         with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
#             # Write each DataFrame to a separate worksheet
#             # div_df.to_excel(writer, sheet_name='Div Data', index=False)
#             table_df.to_excel(writer, sheet_name='Table Data', index=False)

#         print(f"Data has been exported to {excel_file}")
            
#         return True

#     except Exception as e:
#         traceback.print_exc()
#         print("Error occurred while scraping data:", e)
#         return False


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

#     while True:
#         # Process the captcha
#         captcha_text = process_captcha(browser)
#         print("captcha_text", captcha_text)
        
#         # Input the recognized captcha text
#         captcha_input = browser.find_element(By.XPATH, '//*[@id="txt_Captcha"]')
#         captcha_input.send_keys(captcha_text)

#         time.sleep(3)

#         # Click on the element with XPath '//*[@id="viewIEC1"]'
#         view_button = browser.find_element(By.XPATH, '//*[@id="viewIEC1"]')
#         view_button.click()
#         time.sleep(3)

#         # Check if the "Please enter valid captcha code" message is displayed
#         error_message = browser.find_elements(By.XPATH, '//*[@id="incCaptcha"]')
#         if not error_message:
#               soup = BeautifulSoup(browser.page_source, 'html.parser')
              
#               # Print the HTML content to the terminal
#               print(soup.prettify())
#               if scrape_data(browser):
#                     break  # Exit loop if scraping successful

#     else:
#         print("Failed to scrape data from URL:")

# except Exception as e:
#     traceback.print_exc()
#     print("wrong xpath", e)

# finally:
#     browser.quit()


