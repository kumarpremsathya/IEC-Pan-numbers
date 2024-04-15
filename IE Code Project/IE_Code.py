# from selenium import webdriver
# from selenium.webdriver.common.by import By
# import time
# chrome_options = webdriver.ChromeOptions()
# browser = webdriver.Chrome(options=chrome_options)
# browser.get('https://www.dgft.gov.in/CP/?opt=view-any-ice')

# try:
#     next_xpath = browser.find_element(By.XPATH, '//*[@id="mainSectionWrap"]/div[3]/div/div[2]/div[1]')  
#     next_xpath.click()                    
#     time.sleep(2)
#     iec_input = browser.find_element(By.XPATH, '//*[@id="iecNo"]')   #//*[@id="iecNo"]
#     entity_input = browser.find_element(By.XPATH, '//*[@id="entity"]')   #//*[@id="entity"]
#     iec_input.send_keys("12345")
#     time.sleep(3)
#     entity_input.send_keys("magu")
#     time.sleep(3)
#     captcha_input = browser.find_element(By.XPATH, '//*[@id="txt_Captcha"]')      
#     captcha_text = browser.find_element(By.XPATH,  '//*[@id="captcha"]')   
#     captcha_input.send_keys(captcha_text)
#     time.sleep(3)
# except Exception as e:
#     print("wrong xpath", e)
# browser.quit()

# '//img[@id="captcha"]
# //*[@id="txt_Captcha"]



# from selenium import webdriver
# from selenium.webdriver.common.by import By
# import time


# chrome_options = webdriver.ChromeOptions()
# browser = webdriver.Chrome(options=chrome_options)
# browser.get('https://www.dgft.gov.in/CP/?opt=view-any-ice')

# try:
#     next_button = browser.find_element(By.XPATH, '//*[@id="mainSectionWrap"]/div[3]/div/div[2]/div[1]')  
#     next_button.click()                    
#     time.sleep(2)
#     iec_input = browser.find_element(By.XPATH, '//*[@id="iecNo"]')
#     entity_input = browser.find_element(By.XPATH, '//*[@id="entity"]')
#     iec_input.send_keys("12345")
#     time.sleep(3)
#     entity_input.send_keys("magu")
#     time.sleep(3)
    
#     # # Extract and enter captcha text
#     # captcha_text_element = browser.find_element(By.XPATH, '//*[@id="captcha"]')
#     # captcha_input = browser.find_element(By.XPATH, '//*[@id="txt_Captcha"]')
#     # captcha_input.send_keys(captcha_text_element.text)
#     # print("captcha_text_element:", captcha_text_element)
    
#      # Extract and enter captcha text
#     captcha_text_element = browser.find_element(By.XPATH, '//*[@id="captcha"]')
#     print(" captcha_text_element:",  captcha_text_element)
#     captcha_text = captcha_text_element.text  # Extract text from WebElement
#     print("captcha_text:", captcha_text)
#     captcha_input = browser.find_element(By.XPATH, '//*[@id="txt_Captcha"]')
#     captcha_input.send_keys(captcha_text)
#     time.sleep(3)
    
#     time.sleep(3)
# except Exception as e:
#     traceback.print_exc()
#     print("Error:", e)
# finally:
#     browser.quit()
    


# from selenium import webdriver
# from selenium.webdriver.common.by import By
# import time
# from python_anticaptcha import AnticaptchaClient, NoCaptchaTaskProxylessTask
# import traceback
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# import time
# import pytesseract
# from PIL import Image
# import cv2
# import easyocr

# import pytesseract

# pytesseract.pytesseract.tesseract_cmd = r'C:\Users\Premkumar.8265\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'


# chrome_options = webdriver.ChromeOptions()
# browser = webdriver.Chrome(options=chrome_options)
# browser.get('https://www.dgft.gov.in/CP/?opt=view-any-ice')

# try:
#     next_xpath = browser.find_element(By.XPATH, '//*[@id="mainSectionWrap"]/div[3]/div/div[2]/div[1]')
#     next_xpath.click()
#     time.sleep(2)

#     iec_input = browser.find_element(By.XPATH, '//*[@id="iecNo"]')
#     entity_input = browser.find_element(By.XPATH, '//*[@id="entity"]')

#     iec_input.send_keys("12345")
#     time.sleep(3)
#     entity_input.send_keys("magu")
#     time.sleep(3)

#     # Get the captcha image element
#     captcha_image = browser.find_element(By.XPATH, '//*[@id="captcha"]')
    
    
#     # Save the captcha image locally
#     captcha_image.screenshot('captcha.png')
    
#     print("captcha_image : ", captcha_image)
    
#     # Recognize the captcha text using EasyOCR
#     img = cv2.imread('captcha.png')
#     reader = easyocr.Reader(['en'])
#     captcha_text = reader.readtext(img, detail=0)  # Get the recognized text
#     captcha_text = ''.join(captcha_text)  # Convert the list to a string
    
    
#     time.sleep(5)
   
   
    
#     # Input the recognized captcha text
#     captcha_input = browser.find_element(By.XPATH, '//*[@id="txt_Captcha"]')
#     captcha_input.send_keys(captcha_text)

#     time.sleep(3)

# except Exception as e:
    
#     traceback.print_exc()
#     print("wrong xpath", e)
#     browser.quit()
    
    
    
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# import time
# import requests
# import cv2
# import pytesseract

# from selenium import webdriver
# from selenium.webdriver.common.by import By
# import time
# import traceback
# import cv2
# import easyocr
# import sys
# import os
# import time
# import urllib.request


# def progress_hook(blocknum, bs, size, prefix='', suffix='', length=80):
#     """
#     A hook function to show the progress of downloading.
#     """
#     percent = min((blocknum * bs * 100) / size, 100)
#     filled_length = int(length * blocknum * bs // size)
#     bar = '█' * filled_length + '-' * (length - filled_length)  # Use Unicode block character '█'
#     sys.stdout.write('\r{} |{}| {:.1f}% {}'.format(prefix, bar, percent, suffix))  # Use format strings for Unicode characters
#     sys.stdout.flush()


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

#     iec_input.send_keys("12345")
#     time.sleep(3)
#     entity_input.send_keys("magu")
#     time.sleep(3)

#     # Get the captcha image element
#     captcha_image = browser.find_element(By.XPATH, '//*[@id="captcha"]')
    
#     # Save the captcha image locally
#     captcha_image.screenshot('captcha.png')
    
#     # Recognize the captcha text using EasyOCR
#     img = cv2.imread('captcha.png')
#     reader = easyocr.Reader(['en'])
#     captcha_text = reader.readtext(img, detail=0)  # Get the recognized text
#     captcha_text = ''.join(captcha_text)  # Convert the list to a string
    
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

# # Restore stdout
# sys.stdout = sys.__stdout__


import sys
import os
import cv2
import easyocr
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import urllib.request

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

    # Get the captcha image element
    captcha_image = browser.find_element(By.XPATH, '//*[@id="captcha"]')

    # Save the captcha image locally
    with open('captcha.png', 'wb') as file:
        file.write(captcha_image.screenshot_as_png)

    # Redirect stdout to a file for progress output
    with open('output.log', 'w', encoding='utf-8') as f:
        sys.stdout = f
        reader = easyocr.Reader(['en'])
        sys.stdout = sys.__stdout__

    result = reader.readtext('captcha.png', detail=0)
    captcha_text = ''.join(result)

    time.sleep(5)

    # Input the recognized captcha text
    captcha_input = browser.find_element(By.XPATH, '//*[@id="txt_Captcha"]')
    captcha_input.send_keys(captcha_text)

    time.sleep(3)

except Exception as e:
    traceback.print_exc()
    print("wrong xpath", e)

finally:
    browser.quit()

    # Restore stdout
    sys.stdout = sys.__stdout__
    
    

# # 2Captcha API Key
# api_key = "fb32411050ff64ba5228c0908eff4d45"

# chrome_options = webdriver.ChromeOptions()
# browser = webdriver.Chrome(options=chrome_options)
# browser.get('https://www.dgft.gov.in/CP/?opt=view-any-ice')

# try:
#     next_button = browser.find_element(By.XPATH, '//*[@id="mainSectionWrap"]/div[3]/div/div[2]/div[1]')  
#     next_button.click()                    
#     time.sleep(2)
#     iec_input = browser.find_element(By.XPATH, '//*[@id="iecNo"]')
#     entity_input = browser.find_element(By.XPATH, '//*[@id="entity"]')
#     iec_input.send_keys("12345")
#     time.sleep(3)
#     entity_input.send_keys("magu")
#     time.sleep(3)
    
#     # Extract and solve captcha
#     captcha_image_element = browser.find_element(By.XPATH, '//*[@id="captcha"]')
#     captcha_image_base64 = captcha_image_element.get_attribute("src")
#     captcha_text = solve_captcha(api_key, captcha_image_base64)
#     print("captcha_text :", captcha_text)
    
#     # Check if captcha solving was successful
#     if captcha_text is not None:
#         # Enter captcha text
#         captcha_input = browser.find_element(By.XPATH, '//*[@id="txt_Captcha"]')
#         captcha_input.send_keys(captcha_text)
#         time.sleep(3)
#     else:
#        # Prompt user to manually input captcha
#         captcha_text = input("Captcha solving failed. Enter captcha text manually: ")
#         captcha_input = browser.find_element(By.XPATH, '//*[@id="txt_Captcha"]')
#         captcha_input.send_keys(captcha_text)
#         time.sleep(3)
    
# except Exception as e:
#     traceback.print_exc()
#     print("Error:", e)
# finally:
#     browser.quit()





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
