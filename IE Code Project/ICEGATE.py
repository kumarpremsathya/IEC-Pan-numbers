import sys
import os
import cv2
import easyocr
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import urllib.request
from bs4 import BeautifulSoup
import pandas as pd
from selenium.common.exceptions import NoSuchElementException
import json

# # Disable printing progress information during download
# def no_progress(blocknum, bs, size):
#     pass

# # Redirect stdout to a file
# sys.stdout = open(os.devnull, 'w')

# chrome_options = webdriver.ChromeOptions()
# browser = webdriver.Chrome(options=chrome_options)
# browser.get('https://old.icegate.gov.in/EnqMod/')

# try:
#     # next_xpath = browser.find_element(By.XPATH, '//*[@id="mainSectionWrap"]/div[3]/div/div[2]/div[1]')
#     # next_xpath.click()
#     # time.sleep(2)

#     iec_input = browser.find_element(By.XPATH, '//*[@id="searchIECode"]')
    
#     iec_input.send_keys("0388066415")
#     time.sleep(3)
    

#     # Get the captcha image element
#     captcha_image = browser.find_element(By.XPATH, '//*[@id="capimg"]')

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
#     captcha_input = browser.find_element(By.XPATH, '//*[@id="captchaResp"]')
#     captcha_input.send_keys(captcha_text)

#     time.sleep(3)
    

#     # Click on the element with XPath '//*[@id="viewIEC1"]'
#     view_button = browser.find_element(By.XPATH, '//*[@id="pagetable"]/tbody/tr[4]/td[3]/dl/dd/a')
#     view_button.click()
#     time.sleep(15)

# except Exception as e:
#     traceback.print_exc()
#     print("wrong xpath", e)

# finally:
#     browser.quit()

#     # Restore stdout
#     sys.stdout = sys.__stdout__


import sys
import os
import cv2
import easyocr
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import urllib.request

# Function to handle captcha processing
# def process_captcha(browser):
#     # Get the captcha image element
#     captcha_image = browser.find_element(By.XPATH, '//*[@id="capimg"]')

#     # Save the captcha image locally
#     with open('captcha.png', 'wb') as file:
#         file.write(captcha_image.screenshot_as_png)

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
# browser.get('https://old.icegate.gov.in/EnqMod/')

# try:
    
#     while True:
        
#         iec_input = browser.find_element(By.XPATH, '//*[@id="searchIECode"]')    
#         iec_input.send_keys("0388066415")
#         time.sleep(3)
         
#         # Process the captcha
#         captcha_text = process_captcha(browser)

#         # Input the recognized captcha text
#         captcha_input = browser.find_element(By.XPATH, '//*[@id="captchaResp"]')
#         captcha_input.send_keys(captcha_text)

#         time.sleep(3)

#         # Click on the element with XPath '//*[@id="pagetable"]/tbody/tr[4]/td[3]/dl/dd/a'
#         view_button = browser.find_element(By.XPATH, '//*[@id="pagetable"]/tbody/tr[4]/td[3]/dl/dd/a')
#         view_button.click()
#         time.sleep(3)

#         # Check if the "Invalid Code! Please try again!" message is displayed
#         error_message = browser.find_elements(By.XPATH, '//*[@id="pagetable"]/tbody/tr[4]/td[3]/dl/dt[5]/ul/li/span')
#         if not error_message:
#             break  # Break the loop if no error message is found

# except Exception as e:
#     traceback.print_exc()
#     print("wrong xpath", e)

# finally:
#     browser.quit()

#     # Restore stdout
#     sys.stdout = sys.__stdout__





# Function to handle captcha processing
def solve_captcha_icegate1(browser, iec_code):
    try:
        iec_input = browser.find_element(By.XPATH, '//*[@id="searchIECode"]')    
        iec_input.send_keys(iec_code)
        time.sleep(3)
            
        captcha_text = process_captcha_icegate1(browser)
        print("captcha_text " , captcha_text)
        #  Input the recognized captcha text
        
        captcha_input = browser.find_element(By.XPATH, '//*[@id="captchaResp"]')
        captcha_input.send_keys(captcha_text)

        time.sleep(3)

        # Click on the element with XPath '//*[@id="pagetable"]/tbody/tr[4]/td[3]/dl/dd/a'
        view_button = browser.find_element(By.XPATH, '//*[@id="pagetable"]/tbody/tr[4]/td[3]/dl/dd/a')
        view_button.click()
        time.sleep(3)

        # Check if the "Invalid Code! Please try again!" message is displayed
        try:
            error_message = browser.find_element(By.XPATH, '//*[@id="pagetable"]/tbody/tr[4]/td[3]/dl/dt[5]/ul/li/span').text
            print("error_message :", error_message) 
            
        
            if "Invalid Code! Please try again!" in error_message:
                return solve_captcha_icegate1(browser, iec_code)  # Retry solving captcha recursively
                
            else:
                pass
        except Exception as e:
                print("Error in message :", e)
                return True
                
                 
                
    except Exception as e:
        print("Error :", e)
       



# Function to handle captcha processing
def process_captcha_icegate1(browser):
    # Get the captcha image element
    captcha_image = browser.find_element(By.XPATH, '//*[@id="capimg"]')

    # Save the captcha image locally
    with open('captcha.png', 'wb') as file:
        file.write(captcha_image.screenshot_as_png)

    # Load the reader
    reader = easyocr.Reader(['en'])
    result = reader.readtext('captcha.png', detail=0)
    captcha_text = ''.join(result)
    

    return captcha_text

# Disable printing progress information during download
def no_progress(blocknum, bs, size):
    pass

# # Redirect stdout to a file
# sys.stdout = open(os.devnull, 'w')





def scrape_data_icegate1(browser):
    try:
        
    
        # Locate the element containing the data using its XPath
        data_element = browser.find_element(By.XPATH, '//*[@id="sub_content"]/div[2]')
        
        
        # Get the HTML content of the element
        data_html = data_element.get_attribute('innerHTML')
        print("data_html :", data_html)
        
        # Parse the HTML content using BeautifulSoup
        soup = BeautifulSoup(data_html, 'html.parser')
        
                # Find the table
        table = soup.find('table', id='pagetable')
        print("table", table)
        
        
        # Extract the data from the table
        data = {}
        rows = table.find_all('tr')
        address_parts = []
        for row in rows[1:]:  # Skip the header row
            cols = row.find_all('td')
            cols = [col.text.strip() for col in cols]
            if cols:
                field = cols[0].replace('IE Code', 'IEC Number').replace('Name', 'Firm Name')
                value = ' '.join(cols[1:])
                if field in ['Address', '']:  # Combine address parts
                    address_parts.append(value)
                else:
                    data[field] = value

        # Join address parts into a single string
        data['Address'] = ', '.join(address_parts)
        print("data :", data)

        # Create a DataFrame from the extracted data
        df_table1 = pd.DataFrame([data])
        
        
    
        # Parse the HTML content using BeautifulSoup
        soup = BeautifulSoup(data_html, 'html.parser')

       
        
        # Find the second table
        tables = soup.find_all('table', id='pagetable')
        if len(tables) >= 2:
            table = tables[1]
            print("Second table:", table)

            # Extract "Total Number Of Branches" value
            total_branches_row = table.find('th', colspan=True)
            if total_branches_row:
                total_branches_value = total_branches_row.text.split(':')[1].strip()
                print("Total Number Of Branches:", total_branches_value)

            # Extract data from the second table
            branch_details = []
            rows = table.find_all('tr')
            current_branch_detail = {}
            for row in rows:
                cols = row.find_all('td')
                cols = [col.text.strip() for col in cols]
                if cols:
                    if cols[0] == 'Branch Serial Number':
                        if current_branch_detail:
                            branch_details.append(current_branch_detail)
                        current_branch_detail = {'Branch Serial Number': cols[1]}
                    elif cols[0] == 'ADDRESS':
                        current_branch_detail['Address'] = ', '.join(cols[1:])

            # Append the last branch detail
            if current_branch_detail:
                branch_details.append(current_branch_detail)

            print("Branch details:", branch_details)

            # Convert "Branch details" to JSON
            branch_details_json = json.dumps(branch_details)
            
            
            
            # Return the scraped data
            return data, total_branches_value, branch_details_json
        
            # Write "Total Number Of Branches" and "Branch details" to Excel
            # df_table2 = pd.DataFrame({'Total Number Of Branches': [total_branches_value], 'Branch details': [branch_details_json]})
            
            # Concatenate df_table1 and df_table2
            # combined_df = pd.concat([df_table1, df_table2], axis=1)

            # # Write the combined DataFrame to Excel
            # combined_df.to_excel('branch_data.xlsx', index=False)
            # print("Data written to 'branch_data.xlsx'")

            # return True

        else:
            print("Second table not found.")
            return data, None, None

    except Exception as e:
        print("Error occurred while scraping data:", e)
        return None, None, None
    



# Replace 'path/to/your/file.xlsx' with the actual path to your Excel file
df = pd.read_excel('IEC_details.xlsx', dtype={'IEC_CODE': str})


# Create a new DataFrame to store scraped data
scraped_data_df = pd.DataFrame(columns=['IEC Issuance Date', 'DEL Status', 'IEC Number', 'Firm Name', 'Address', 'Total Number Of Branches', 'Branch details'])




def icegate_first(iec_code):
    global scraped_data_df # Add this line to access the global variable
    chrome_options = webdriver.ChromeOptions()
    
    
    # Add preferences to clear cache
    # chrome_prefs = {}
    # chrome_prefs["profile.default_content_settings"] = {"images": 2, "plugins": 2, "popups": 2, "geolocation": 2, 
    #                                                     "notifications": 2, "auto_select_certificate": 2, "fullscreen": 2, 
    #                                                     "mouselock": 2, "mixed_script": 2, "media_stream": 2, 
    #                                                     "media_stream_mic": 2, "media_stream_camera": 2, "protocol_handlers": 2,
    #                                                     "ppapi_broker": 2, "automatic_downloads": 2, "midi_sysex": 2, 
    #                                                     "push_messaging": 2, "ssl_cert_decisions": 2, "metro_switch_to_desktop": 2, 
    #                                                     "protected_media_identifier": 2, "app_banner": 2, "site_engagement": 2, 
    #                                                     "durable_storage": 2}
    # chrome_prefs["profile.default_content_setting_values"] = {"cookies": 2, "images": 2, "javascript": 1, "plugins": 2, "popups": 2, 
    #                                                           "geolocation": 2, "notifications": 2, "auto_select_certificate": 2, 
    #                                                           "fullscreen": 2, "mouselock": 2, "mixed_script": 2, "media_stream": 2, 
    #                                                           "media_stream_mic": 2, "media_stream_camera": 2, "protocol_handlers": 2,
    #                                                           "ppapi_broker": 2, "automatic_downloads": 2, "midi_sysex": 2, 
    #                                                           "push_messaging": 2, "ssl_cert_decisions": 2, "metro_switch_to_desktop": 2, 
    #                                                           "protected_media_identifier": 2, "app_banner": 2, "site_engagement": 2, 
    #                                                           "durable_storage": 2}
    # chrome_options.add_experimental_option("prefs", chrome_prefs)
    
    browser = webdriver.Chrome(options=chrome_options)
    browser.get('https://old.icegate.gov.in/EnqMod/')

    try:
        
   
            
        iec_input = browser.find_element(By.XPATH, '//*[@id="searchIECode"]')    
        iec_input.send_keys(iec_code)
        time.sleep(3)
        
    
    
          # Attempt to solve the captcha
        try:
            captcha_solved_icegate1 = solve_captcha_icegate1(browser, iec_code)
            print("captcha_solved_icegate1:", captcha_solved_icegate1)    
        
        except Exception as e:
            traceback.print_exc()
            print("captcha_solved:", e)    
        
        time.sleep(5)

        if captcha_solved_icegate1:
        # Wait for a few seconds before checking for the presence of the message
            time.sleep(3)
            
             # Check if the message "Details for this IEC Number is not available." exists
            try:
                unavailable_message = browser.find_element(By.XPATH, '//*[@id="sub_content"]/div[2]').text
                if "No Record Found" in unavailable_message:
                    # icegate_second()  # Call the icegate function if the message exists
                    pass
                else:
                    # Proceed with scraping the data if the message doesn't exist
                    # Define the column names to be set as NULL
                    null_columns = [
                        "IEC Issuance Date", "IEC Status", "DEL Status", "IEC Cancelled Date",
                        "IEC Suspended Date", "File Number", "File Date", "DGFT RA Office", "Category of Exporters",
                        "RCMC DETAILS", "Nature of concern/Firm"
                    ]
                    
                    # Create a dictionary with NULL values for the specified columns
                    null_data = {column: ["NULL"] for column in null_columns}
                    
                    
                    scraped_data = scrape_data_icegate1(browser)
                    if scraped_data is not None:
                        data, total_branches_value, branch_details_json = scraped_data
                        # Create a new row in the DataFrame with the scraped data
                        new_row = {
                            'IEC Issuance Date': null_data.get('IEC Issuance Date',''),
                            'DEL Status': null_data.get('DEL Status',''),
                            'IEC Number': data.get('IEC Number', ''),
                            'Firm Name': data.get('Firm Name', ''),
                            'Address': data.get('Address', ''),
                            'Total Number Of Branches': total_branches_value,
                            'Branch details': branch_details_json
                        }
                        
                        scraped_data_df= pd.concat([scraped_data_df, pd.DataFrame([new_row])], ignore_index=True)
                        # Write the scraped data to an Excel file
                        # scraped_data.to_excel('scraped_data.xlsx', index=False)
                        
                   
                        
            except NoSuchElementException:
                # If the element is not found, proceed with scraping the data
                # scrape_data_icegate1(browser)
                pass
            
    except Exception as e:
        traceback.print_exc()
        print("wrong xpath", e)

    finally:
        browser.quit()

        # Restore stdout
        sys.stdout = sys.__stdout__




def read_excel():
    try:
        for index, row in df.head(2).iterrows(): 
            iec_code = row['IEC_CODE']
            print("iec_code", iec_code)
            # firm_name = row['FIRM NAME']
            
            # Call your function or code with the IEC code and firm name
            icegate_first(iec_code)
    except Exception as e:
        traceback.print_exc()
        print("Excel reading  error", e)
        
     # Write the scraped data to an Excel file
    scraped_data_df.to_excel('scraped_data.xlsx', index=False)
    
read_excel()