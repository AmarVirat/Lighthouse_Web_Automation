from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui as pg
import sys,six
import urllib.request
import ssl
from datetime import date

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import pandas as pd
from datetime import datetime, timedelta
import psycopg2 as ps
import shutil
import os

def get_downloaded_data(source_folder, destination_folder, pdf_filename):
    """
    Moves a PDF file from source_folder to destination_folder.

    Parameters:
    - source_folder: The path to the source folder.
    - destination_folder: The path to the destination folder.
    - pdf_filename: The name of the PDF file to move.
    """
    # Construct full paths
    source_path = os.path.join(source_folder, pdf_filename)
    destination_path = os.path.join(destination_folder, pdf_filename)

    print("=========== Source Path ===============>", source_path)
    print("=========== Destination Path ===============>", destination_path)
    print("=======Pdf File Name ==================>", pdf_filename)

    # Check if the source file exists
    if os.path.exists(source_path):
        try:
            # Move the PDF file
            shutil.move(source_path, destination_path)
            print(f"'{pdf_filename}' has been successfully moved from {source_folder} to {destination_folder}.")
        except Exception as e:
            print(f"Error moving file: {e}")
    else:
        print(f"The file '{pdf_filename}' does not exist in the source folder.")

def update_tracker(Status_sucess_failure, logging_str):
    # Define the file path
    print("================I am in Update Tracker ==================================>")
    file_path = r'C:\Users\Administrator\OneDrive - Indorama Ventures PCL\Project LightHouse\Historical Data\CMA_Other_scripts\Tracker.xlsx'

    # Create sample data to update
    new_data = {
        'filename': ['Woodmac_test_monthly.py'],
        'frequency': ['Monthly'],
        'status': [Status_sucess_failure],
        'date': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')], # Current timestamp
        'Logs' : [logging_str]
    }

    new_df = pd.DataFrame(new_data)

    # Check if the CSV file already exists
    if os.path.exists(file_path):
        existing_df = pd.read_excel(file_path)
        updated_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        updated_df = new_df

    updated_df.to_excel(file_path, index=False)
    return("Succesfully Updated")


def PTA_Shortterm_Monthly():
    logs_str = ''
    try:
        logs_str = ''+ logs_str +'\n'+'Entering PTA Shortterm '
        chrome_options = Options()
        ### Suppress printer to switch to Save As option
        chrome_options.add_argument('--kiosk-printing')
        driver_path = r"C:\Users\Administrator\Downloads\chromedriver-win64\chromedriver-win64\chromedriver.exe"

        service=Service(executable_path=driver_path)
        driver = webdriver.Chrome(options=chrome_options)
        driver.maximize_window()

        driver.get('https://my.woodmac.com/')


        time.sleep(3)

        driver.find_element(By.XPATH,"//*[@id=\"onetrust-reject-all-handler\"]").click()
        time.sleep(3)

        username_field = driver.find_element(By.XPATH,"//*[@id=\"idp-discovery-username\"]")
        time.sleep(3)
        #password_field = driver.find_element(By.ID, 'password')  # Adjust selector as needed
        print("=========@IVLstrategy2024======> Logging in")
        username_field.send_keys('pongsatorn.v@indorama.net')

        time.sleep(3)
        logs_str = ''+ logs_str +'\n'+'Entering PTA Shortterm :: Logged In Complete'
        driver.find_element(By.XPATH, "//*[@id=\"idp-discovery-submit\"]").click()
        time.sleep(3)

        password_field = driver.find_element(By.XPATH,"//*[@id=\"okta-signin-password\"]")
        password_field.send_keys('#IVLstrategy2025')

        time.sleep(3)

        driver.find_element(By.XPATH, "//*[@id=\"okta-signin-submit\"]").click()
        time.sleep(3)

        driver.find_element(By.XPATH, "/html/body/header/div/button").click()

        Search_box = driver.find_element(By.XPATH, "//*[@id=\"globalPredictiveInput\"]")
        Search_box.click()

        Search_box.send_keys("Paraxylene & PTA Short Term (Monthly)")   
        
        pg.press('enter')

        time.sleep(5)

        search_pdf = 'Paraxylene and derivatives global monthly market overview '+datetime.now().strftime("%B %Y")
        logs_str = ''+ logs_str +'\n'+'Entering PTA Shortterm :: Searching PDF'
        print(search_pdf)
        # Wait for the page to load
        time.sleep(2)
        

        # Infinite scroll loop to scroll until the link is found
        while True:
            try:
                # Find the link element by its partial text
                link = driver.find_element(By.PARTIAL_LINK_TEXT, search_pdf)
                
                # Click the link
                link.click()
                break  # Exit the loop once the link is clicked
            except Exception as e:
                # If the link is not found, scroll down and try again
                pdf_found = False
                print("Link not found, scrolling down...")
                driver.execute_script("window.scrollBy(0, 50);")  # Scroll down by 1000 pixels
                time.sleep(2)  # Wait for the page to load the new content


        logs_str = ''+ logs_str +'\n'+'Entering PTA Shortterm :: Found the link'
        time.sleep(5)

        pdf_found = False
        while True:
            try:
                pdf_name = str(str(search_pdf).title())+".pdf"
                print(pdf_name)
                download_pdf = driver.find_element(By.PARTIAL_LINK_TEXT, pdf_name)
                download_pdf.click()
                pdf_found = True
            
                break
            except Exception as e:
                pdf_found = False
                print(pdf_name)
                print("Link not found, scrolling down...")
                driver.execute_script("window.scrollBy(0, 50);")  # Scroll down by 1000 pixels
                time.sleep(2)  # Wait for the page to load the new content

        time.sleep(6)
        driver.quit()
        logs_str = ''+ logs_str +'\n'+'Entering PTA Shortterm :: Found the pdf'
        source_folder = r"C:\Users\Administrator\Downloads"
        destination_folder =  r"C:\Users\Administrator\OneDrive - Indorama Ventures PCL\01 Industry Report\WM\Paraxylene & PTA Short Term (Monthly)"+'\\'+str(datetime.now().strftime('%Y'))
        logs_str = ''+ logs_str +'\n'+'Entering PTA Shortterm '

        if(pdf_found == True):
            df_cur = get_downloaded_data(source_folder, destination_folder, str(pdf_name.lower()).replace(" ","-"))
            logs_str = ''+ logs_str +'\n'+'Entering PTA Shortterm :: Moved PDF to Destination'
            updated_status = update_tracker("Success", logs_str)
            print(updated_status)
        else:
            logs_str = ''+ logs_str +'\n'+'Entering PTA Shortterm :: Failed Moved PDF to Destination'
            updated_status = update_tracker("Failure", logs_str)
            print(updated_status)
    except Exception as e:
        logs_str = logs_str + '\n'+ "Failed Downloading PDF ::"+str(e)
        updated_status = update_tracker("Failure", logs_str)
        print(updated_status)

def polyster_raw_materials():
    logs_str = ''
    try:
        logs_str = ''+ logs_str +'\n'+'Entering Polyster Raw Materials Shortterm '
        chrome_options = Options()
        ### Suppress printer to switch to Save As option
        chrome_options.add_argument('--kiosk-printing')
        driver_path = r"C:\Users\Administrator\Downloads\chromedriver-win64\chromedriver-win64\chromedriver.exe"

        service=Service(executable_path=driver_path)
        driver = webdriver.Chrome(options=chrome_options)
        driver.maximize_window()

        driver.get('https://my.woodmac.com/')


        time.sleep(3)

        driver.find_element(By.XPATH,"//*[@id=\"onetrust-reject-all-handler\"]").click()
        time.sleep(3)

        username_field = driver.find_element(By.XPATH,"//*[@id=\"idp-discovery-username\"]")
        time.sleep(3)
        #password_field = driver.find_element(By.ID, 'password')  # Adjust selector as needed
        print("=========@IVLstrategy2024======> Logging in")
        username_field.send_keys('pongsatorn.v@indorama.net')

        time.sleep(3)

        driver.find_element(By.XPATH, "//*[@id=\"idp-discovery-submit\"]").click()
        time.sleep(3)

        password_field = driver.find_element(By.XPATH,"//*[@id=\"okta-signin-password\"]")
        password_field.send_keys('@IVLstrategy2024')

        time.sleep(3)

        driver.find_element(By.XPATH, "//*[@id=\"okta-signin-submit\"]").click()
        time.sleep(3)
        logs_str = ''+ logs_str +'\n'+'Login completed '

        driver.find_element(By.XPATH, "/html/body/header/div/button").click()

        Search_box = driver.find_element(By.XPATH, "//*[@id=\"globalPredictiveInput\"]")
        Search_box.click()

        Search_box.send_keys("Polyester & Raw Materials Short Term (Monthly)")

        pg.press('enter')
        logs_str = ''+ logs_str +'\n'+'Entering polyster_raw_materials:: Searching PDF'
        time.sleep(5)

        search_pdf = 'Polyester and raw materials global monthly market overview '+datetime.now().strftime("%B %Y")

        print(search_pdf)
        # Wait for the page to load
        time.sleep(2)

        # Infinite scroll loop to scroll until the link is found
        while True:
            try:
                # Find the link element by its partial text
                link = driver.find_element(By.PARTIAL_LINK_TEXT, search_pdf)
                
                # Click the link
                link.click()
                break  # Exit the loop once the link is clicked
            except Exception as e:
                pdf_found = True
                # If the link is not found, scroll down and try again
                print("Link not found, scrolling down...")
                driver.execute_script("window.scrollBy(0, 50);")  # Scroll down by 1000 pixels
                time.sleep(2)  # Wait for the page to load the new content

        
        logs_str = ''+ logs_str +'\n'+'Entering polyster_raw_materials :: Found the link'

        time.sleep(5)
        pdf_found = False
        while True:
            try:
                pdf_name = str(str(search_pdf).title())+".pdf"
                print(pdf_name)
                download_pdf = driver.find_element(By.PARTIAL_LINK_TEXT, pdf_name)
                download_pdf.click()
                pdf_found = True
                break
            except Exception as e:
                print(pdf_name)
                print("Link not found, scrolling down...")
                driver.execute_script("window.scrollBy(0, 50);")  # Scroll down by 1000 pixels
                time.sleep(2)  # Wait for the page to load the new content
                pdf_found = False
        time.sleep(6)

        driver.quit()

        logs_str = ''+ logs_str +'\n'+'Entering polyster_raw_materials :: Found the pdf'
    
        source_folder = r"C:\Users\Administrator\Downloads"
        destination_folder = r"C:\Users\Administrator\OneDrive - Indorama Ventures PCL\01 Industry Report\WM\Polyester & Raw Materials Short Term (Monthly)"+'\\'+str(datetime.now().strftime('%Y'))
        if(pdf_found == True):
            logs_str = ''+ logs_str +'\n'+'Entering PTA Shortterm :: Moving PDF to Destination'
            df_cur = get_downloaded_data(source_folder, destination_folder, str(pdf_name.lower()).replace(" ","-"))
            logs_str = ''+ logs_str +'\n'+'Entering polyster_raw_materials:: Moved PDF to Destination'
            updated_status = update_tracker("Success")
            print(updated_status)
        else:
            logs_str = ''+ logs_str +'\n'+'Entering polyster_raw_materials:: Failed Moved PDF to Destination'
            updated_status = update_tracker("Failure")
            print(updated_status)
    except Exception as e:
        logs_str = logs_str + '\n'+ "Failed Downloading PDF ::"+str(e)
        updated_status = update_tracker("Failure", logs_str)
        print(updated_status)


def Ethylene_oxide():
    logs_str = ''
    logs_str = ''+ logs_str +'\n'+'Entering Ethylene_oxide :: Logged In Complete'

    try:
        chrome_options = Options()
        ### Suppress printer to switch to Save As option
        chrome_options.add_argument('--kiosk-printing')
        driver_path = r"C:\Users\Administrator\Downloads\chromedriver-win64\chromedriver-win64\chromedriver.exe"

        service=Service(executable_path=driver_path)
        driver = webdriver.Chrome(options=chrome_options)
        driver.maximize_window()

        driver.get('https://my.woodmac.com/')


        time.sleep(3)

        driver.find_element(By.XPATH,"//*[@id=\"onetrust-reject-all-handler\"]").click()
        time.sleep(3)

        username_field = driver.find_element(By.XPATH,"//*[@id=\"idp-discovery-username\"]")
        time.sleep(3)
        #password_field = driver.find_element(By.ID, 'password')  # Adjust selector as needed
        print("=========@IVLstrategy2024======> Logging in")
        username_field.send_keys('pongsatorn.v@indorama.net')

        time.sleep(3)

        driver.find_element(By.XPATH, "//*[@id=\"idp-discovery-submit\"]").click()
        time.sleep(3)

        password_field = driver.find_element(By.XPATH,"//*[@id=\"okta-signin-password\"]")
        password_field.send_keys('@IVLstrategy2024')

        time.sleep(3)

        driver.find_element(By.XPATH, "//*[@id=\"okta-signin-submit\"]").click()
        time.sleep(3)
        logs_str = ''+ logs_str +'\n'+'Entering Ethylene_oxide :: Searching PDF'


        driver.find_element(By.XPATH, "/html/body/header/div/button").click()

        Search_box = driver.find_element(By.XPATH, "//*[@id=\"globalPredictiveInput\"]")
        Search_box.click()

        Search_box.send_keys("Ethylene Oxide & Glycol Short Term (Monthly)")

        pg.press('enter')

        time.sleep(5)

        search_pdf = 'Ethylene oxide and glycol global monthly market overview '+datetime.now().strftime("%B %Y")

        print(search_pdf)
        # Wait for the page to load
        time.sleep(2)

        # Infinite scroll loop to scroll until the link is found
        while True:
            try:
                # Find the link element by its partial text
                link = driver.find_element(By.PARTIAL_LINK_TEXT, search_pdf)
                
                # Click the link
                link.click()
                logs_str = ''+ logs_str +'\n'+'Entering Ethylene_oxide:: Found the link'

                break  # Exit the loop once the link is clicked
            except Exception as e:
                pdf_found = False
                # If the link is not found, scroll down and try again
                print("Link not found, scrolling down...")
                driver.execute_script("window.scrollBy(0, 50);")  # Scroll down by 1000 pixels
                time.sleep(2)  # Wait for the page to load the new content

        pdf_found = False
        time.sleep(5)

        while True:
            try:
                pdf_name = str(str(search_pdf).title())+".pdf"
                print(pdf_name)
                download_pdf = driver.find_element(By.PARTIAL_LINK_TEXT, pdf_name)
                download_pdf.click()
                pdf_found == True
                logs_str = ''+ logs_str +'\n'+'Entering Ethylene_oxide :: Found PDF'
                break
            except Exception as e:
                pdf_found = False
                print(pdf_name)
                print("Link not found, scrolling down...")
                driver.execute_script("window.scrollBy(0, 50);")  # Scroll down by 1000 pixels
                time.sleep(2)  # Wait for the page to load the new content
        time.sleep(6)

        driver.quit()                                                                        

        source_folder = r"C:\Users\Administrator\Downloads"
        destination_folder = r"C:\Users\Administrator\OneDrive - Indorama Ventures PCL\01 Industry Report\WM\Ethylene Oxide & Glycol Short Term (Monthly)"+'\\'+str(datetime.now().strftime('%Y'))
        if(pdf_found == True):
            logs_str = ''+ logs_str +'\n'+'Entering Ethylene_oxide :: Moved PDF to Destination'
            df_cur = get_downloaded_data(source_folder, destination_folder, str(pdf_name.lower()).replace(" ","-"))
            updated_status = update_tracker("Success")
            print(updated_status)
        else:
            logs_str = ''+ logs_str +'\n'+'Entering Ethylene_oxide :: Dint Move PDF to Destination'
            updated_status = update_tracker("Failure")
            print(updated_status)
    except Exception as e:
        logs_str = logs_str + '\n'+ "Failed Downloading PDF ::"+str(e)
        updated_status = update_tracker("Failure", logs_str)


PTA_Shortterm_Monthly()
polyster_raw_materials()
Ethylene_oxide()