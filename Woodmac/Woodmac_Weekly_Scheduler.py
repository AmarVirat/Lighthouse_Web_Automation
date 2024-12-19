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
        'filename': ['Woodmac_Weekly_Scheduler.py'],
        'frequency': ['Weekly'],
        'status': [Status_sucess_failure],
        'date': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
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


def Polyester_and_Raw_Materials_Short_Term_Weekly():
    try:
        logs_str = ' '

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

        logs_str = ' '+ 'Successfully Logged In '+'\n'


        driver.find_element(By.XPATH, "/html/body/header/div/button").click()

        Search_box = driver.find_element(By.XPATH, "//*[@id=\"globalPredictiveInput\"]")
        Search_box.click()

        Search_box.send_keys("Polyester & Raw Materials Short Term (Weekly)")

        pg.press('enter')

        time.sleep(5)

        # Polyester and raw materials weekly market overview 2 December 2024

        search_pdf_list = []
        for date in range(7):

            # Get the current date as a string with leading zero
            date_str = datetime.now() - timedelta(days=date) 

            date_str = date_str.strftime("%d %B %Y") 
            # Remove the leading zero from the day (if it exists)
            day, month, year = date_str.split(' ')
            day = day.lstrip('0')  # Remove leading zero from the day

            # Reassemble the date string without the leading zero
            formatted_date = f"{day} {month} {year}"
            search_pdf_list.append(formatted_date)

        print(search_pdf_list)
        # Wait for the page to load
        time.sleep(2)

        last_scroll_position = 0
        scroll_attempts = 0
        max_scroll_attempts = 5  # Limit the number of scroll attempts to avoid infinite scrolling

        # Infinite scroll loop to scroll until the link is found
        while True:
            try:
                # Check if we've scrolled too many times without finding a link
                if scroll_attempts > max_scroll_attempts:
                    print("Reached max scroll attempts, exiting.")
                    break

                # Search for the link in search_pdf_list
                link_found = False
                for search_pdf in search_pdf_list:
                    try:
                        search_pdf_link = 'Polyester and raw materials weekly market overview '+search_pdf
                        link = driver.find_element(By.PARTIAL_LINK_TEXT, search_pdf_link)
                        print(f"Link found: {search_pdf_link}")
                        link.click()  # Click the link
                        link_found = True
                        break  # Exit the loop once the link is clicked
                    except Exception as e:
                        # Link not found for this term, continue with the next search term
                        continue

                # If a link was found, break the outer loop
                if link_found:
                    break

                # If no link was found, scroll the page
                print("Link not found, scrolling down...")
                driver.execute_script("window.scrollBy(0, 500);")  # Scroll down by 500 pixels
                time.sleep(2)  # Wait for the page to load the new content

                # Check if the scroll position has changed to detect if we've reached the end of the page
                current_scroll_position = driver.execute_script("return window.pageYOffset;")
                if current_scroll_position == last_scroll_position:
                    print("No new content loaded after scrolling, exiting.")
                    break  # Exit the loop if no new content loaded

                # Update the last scroll position
                last_scroll_position = current_scroll_position
                scroll_attempts += 1

            except Exception as e:
                print(f"An error occurred: {e}")
                break  # Exit the loop if an unexpected error occurs

        time.sleep(5)

        logs_str = ' '+ 'PDF Linked found for Polyester and raw materials weekly market overview'+'\n'



        while True:
            try:
                pdf_name = str(str("Polyester Raw Materials Global Weekly "+str(search_pdf)).title())+".pdf"
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
        logs_str = ' '+ 'PDF Downloaded Successfully for Polyester and raw materials weekly market overview'+'\n'
        time.sleep(6)

        driver.quit()

        source_folder = r"C:\Users\Administrator\Downloads"
        destination_folder =  r"C:\Users\Administrator\OneDrive - Indorama Ventures PCL\01 Industry Report\WM\Polyester & Raw Materials Short Term (Weekly)"+'\\'+str(datetime.now().strftime('%Y'))

        if(pdf_found == True):
            df_cur = get_downloaded_data(source_folder, destination_folder, str(pdf_name.lower()).replace(" ","-"))
            logs_str = ''+ logs_str +'\n'+'Entering  Polyester and raw materials weekly market overview :: Moved PDF to Destination'
            updated_status = update_tracker("Success", logs_str)
            print(updated_status)
        else:
            logs_str = ''+ logs_str +'\n'+'Entering  Polyester and raw materials weekly market overview :: Failed Moved PDF to Destination'
            updated_status = update_tracker("Failure", logs_str)
            print(updated_status)
    except Exception as e:    
        logs_str = ' '+ logs_str +'\n'+'f"An error occurred: {e}'
        updated_status = update_tracker("Failure", logs_str)



def Polyester_and_Raw_Materials_Asia_Daily_Express_Weekly():
    try:
        logs_str = ''
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

        logs_str = ' '+ 'Successfully Logged In '+'\n'

        driver.find_element(By.XPATH, "/html/body/header/div/button").click()

        Search_box = driver.find_element(By.XPATH, "//*[@id=\"globalPredictiveInput\"]")
        Search_box.click()

        Search_box.send_keys("Polyester & Raw Materials Asia Daily Express (Weekly)")

        pg.press('enter')

        time.sleep(5)

        search_pdf = 'Polyester and raw materials Asia daily express '+datetime.now().strftime("%B %Y")

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
                logs_str = ' '+ 'PDF Linked found for Polyester & Raw Materials Asia Daily Express (Weekly) market overview'+'\n'
                break  # Exit the loop once the link is clicked
            except Exception as e:
                # If the link is not found, scroll down and try again
                print("Link not found, scrolling down...")
                driver.execute_script("window.scrollBy(0, 50);")  # Scroll down by 1000 pixels
                time.sleep(2)  # Wait for the page to load the new content

        
        time.sleep(5)



        while True:
            try:
                # pdf_name = str(str(search_pdf).title())+".pdf"     ###   'Polyester and raw materials Asia daily express '+datetime.now().strftime("%B %Y")  Might change
                pdf_name = search_pdf
                download_pdf = driver.find_element(By.PARTIAL_LINK_TEXT, pdf_name)
                print(pdf_name+"Found")
                download_pdf.click()
                pdf_found = True
                logs_str = ' '+ 'PDF Downloaded Successfully for Polyester & Raw Materials Asia Daily Express (Weekly) market overview'+'\n'
                break
            except Exception as e:
                print(pdf_name)
                pdf_found = False
                print("Link not found, scrolling down...")
                driver.execute_script("window.scrollBy(0, 50);")  # Scroll down by 1000 pixels
                time.sleep(2)  # Wait for the page to load the new content

        time.sleep(6)

        driver.quit()

        


        logs_str = ''+ logs_str +'\n'+'Entering Polyester & Raw Materials Asia Daily Express (Weekly) Shortterm :: Found the pdf'
        source_folder = r"C:\Users\Administrator\Downloads"
        destination_folder =  r"C:\Users\Administrator\OneDrive - Indorama Ventures PCL\01 Industry Report\WM\Polyester & Raw Materials Asia Daily Express (Weekly)"+'\\'+str(datetime.now().strftime('%Y'))

        if(pdf_found == True):
            df_cur = get_downloaded_data(source_folder, destination_folder, str(pdf_name.lower()).replace(" ","-"))
            logs_str = ''+ logs_str +'\n'+'Entering Polyester & Raw Materials Asia Daily Express (Weekly) :: Moved PDF to Destination'
            updated_status = update_tracker("Success", logs_str)
            print(updated_status)
        else:
            logs_str = ''+ logs_str +'\n'+'Entering Polyester & Raw Materials Asia Daily Express (Weekly) :: Failed Moved PDF to Destination'
            updated_status = update_tracker("Failure", logs_str)
            print(updated_status)
    except Exception as e:    
        logs_str = ' '+ logs_str +'\n'+'f"An error occurred: {e}'
        updated_status = update_tracker("Failure", logs_str)





def PET_Asia_Market_Overview_Weekly():
    logs_str = ''
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

        logs_str = ' '+ 'Successfully Logged In '+'\n'


        driver.find_element(By.XPATH, "/html/body/header/div/button").click()

        Search_box = driver.find_element(By.XPATH, "//*[@id=\"globalPredictiveInput\"]")
        Search_box.click()

        Search_box.send_keys("PET Asia Market Overview (Weekly)")

        pg.press('enter')

        time.sleep(5)
        search_pdf_list = []
        for date in range(7):

            # Get the current date as a string with leading zero
            date_str = datetime.now() - timedelta(days=date) 

            date_str = date_str.strftime("%d %B %Y") 

            # Remove the leading zero from the day (if it exists)
            day, month, year = date_str.split(' ')
            day = day.lstrip('0')  # Remove leading zero from the day

            # Reassemble the date string without the leading zero
            formatted_date = f"{day} {month} {year}"
            search_pdf_list.append(formatted_date)

        print(search_pdf_list)
        # Wait for the page to load
        time.sleep(2)


        last_scroll_position = 0
        scroll_attempts = 0
        max_scroll_attempts = 7  # Limit the number of scroll attempts to avoid infinite scrolling

        # Infinite scroll loop to scroll until the link is found
        while True:
            try:
                # Check if we've scrolled too many times without finding a link
                if scroll_attempts > max_scroll_attempts:
                    print("Reached max scroll attempts, exiting.")
                    break

                # Search for the link in search_pdf_list
                link_found = False
                for search_pdf in search_pdf_list:
                    try:
                        search_pdf_link = 'PET Asia weekly market overview '+search_pdf
                        print("Searching For============>",search_pdf_link)
                        link = driver.find_element(By.PARTIAL_LINK_TEXT, search_pdf_link)
                        print(f"Link found: {search_pdf_link}")
                        link.click()  # Click the link
                        link_found = True
                        logs_str = ' '+ 'PDF Linked found for PET Asia Market Overview (Weekly) market overview'+'\n'
                        break  # Exit the loop once the link is clicked
                    except Exception as e:
                        # Link not found for this term, continue with the next search term
                        continue

                # If a link was found, break the outer loop
                if link_found:
                    break

                # If no link was found, scroll the page
                print("Link not found, scrolling down...")
                driver.execute_script("window.scrollBy(0, 500);")  # Scroll down by 500 pixels
                time.sleep(2)  # Wait for the page to load the new content

                # Check if the scroll position has changed to detect if we've reached the end of the page
                current_scroll_position = driver.execute_script("return window.pageYOffset;")
                if current_scroll_position == last_scroll_position:
                    print("No new content loaded after scrolling, exiting.")
                    break  # Exit the loop if no new content loaded

                # Update the last scroll position
                last_scroll_position = current_scroll_position
                scroll_attempts += 1

            except Exception as e:
                print(f"An error occurred: {e}")
                break  # Exit the loop if an unexpected error occurs

        time.sleep(5)

        


        while True:
            try:
                # pdf_name = str(str("Polyester Raw Materials Global Weekly "+str(search_pdf)).title())+".pdf"
                pdf_name = search_pdf_link
                print(pdf_name)
                download_pdf = driver.find_element(By.PARTIAL_LINK_TEXT, pdf_name)
                download_pdf.click()
                pdf_found = True
                logs_str = ' '+ 'PDF Downloaded Successfully for PET Asia Market Overview (Weekly) market overview'+'\n'
                break
            except Exception as e:
                pdf_found = False
                print(pdf_name)
                print("Link not found, scrolling down...")
                driver.execute_script("window.scrollBy(0, 50);")  # Scroll down by 1000 pixels
                time.sleep(2)  # Wait for the page to load the new content

        time.sleep(10)


        driver.quit()


        logs_str = ''+ logs_str +'\n'+'Entering PTA Shortterm :: Found the pdf'
        source_folder = r"C:\Users\Administrator\Downloads"
        destination_folder =  r"C:\Users\Administrator\OneDrive - Indorama Ventures PCL\01 Industry Report\WM\PET Asia Market Overview (Weekly)"+'\\'+str(datetime.now().strftime('%Y'))

        if(pdf_found == True):
            df_cur = get_downloaded_data(source_folder, destination_folder, str(pdf_name.lower()).replace(" ","-"))
            logs_str = ''+ logs_str +'\n'+'Entering PET Asia Market Overview (Weekly) :: Moved PDF to Destination'
            updated_status = update_tracker("Success", logs_str)
            print(updated_status)
        else:
            logs_str = ''+ logs_str +'\n'+'Entering PET Asia Market Overview (Weekly) :: Failed Moved PDF to Destination'
            updated_status = update_tracker("Failure", logs_str)
            print(updated_status)
    except Exception as e:    
        logs_str = ' '+ logs_str +'\n'+'f"An error occurred: {e}'
        updated_status = update_tracker("Failure", logs_str)



Polyester_and_Raw_Materials_Short_Term_Weekly()
Polyester_and_Raw_Materials_Asia_Daily_Express_Weekly()
PET_Asia_Market_Overview_Weekly()