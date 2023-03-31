from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
from bs4 import BeautifulSoup
import time
import pandas as pd
import math
options = webdriver.ChromeOptions()
options.add_argument('log-level=3')
options.add_argument('--ignore-certificate-errors-spki-list')
options.add_argument('--ignore-ssl-errors')
# options.add_argument('--headless')
options.add_argument('--disable-gpu')
s = Service('D:\web_scraping_project\chromedriver.exe')
driver = webdriver.Chrome(service=s, options=options)
driver.maximize_window()
url = "https://ipindiaservices.gov.in/publicsearch"
driver.get(url)
print(driver.title)

driver.find_element(By.ID, "FromDate").send_keys("03-14-2023")
driver.find_element(By.ID, "ToDate").send_keys("03-14-2023")

captcha = driver.find_element(By.XPATH, '//*[@id="CaptchaText"]')
driver.execute_script("arguments[0].scrollIntoView();", captcha)
captcha.click()

time.sleep(7)

doc = driver.find_element(
    By.XPATH, '//*[@id="header"]/div[4]/div/div[1]/div[2]')
doc_str = doc.text
temp_list = list(doc_str.split())
pages = int(temp_list[-1])
pages_click = (math.ceil(pages / 25)) - 1
print("Total no. of pages = ", pages_click + 1)
print("pages click:", pages_click)
print("Each page has 25 rows.")
print("Total rows:", pages)
master_df = pd.DataFrame(columns=[
    'Application Number', 'Title', 'Application Date', 'Status', 'Publication Number', 'Publication Date(U/S 11A)', 'Publication Type', 'Application Filing Date', 'Priority Number', 'Priority Country', 'Priority Date', 'Field Of Invention', 'Classification (IPC)', 'Inventor Name', 'Inventor Address', 'Inventor Country', 'Inventor Nationality', 'Applicant Name', 'Applicant Address', 'Applicant Country', 'Applicant Nationality', 'Application Type', 'E-MAIL (As Per Record)', 'ADDITIONAL-EMAIL (As Per Record)', 'E-MAIL (UPDATED Online)', 'REQUEST FOR EXAMINATION DATE', 'Application Status'])

for i in range(pages_click + 1):
    print(f"Page Number : {i + 1}")
    # Data from the search page.
    Application_Number = []
    Title = []
    Application_Date = []
    Status = []

    # Data from Application Number link.
    publication_number2 = []
    publication_date2 = []
    publication_type2 = []
    application_filing_date2 = []
    priority_number2 = []
    priority_country2 = []
    priority_date2 = []
    field_of_invention2 = []
    classification_ipc2 = []

    inventor_name = []
    inventor_address = []
    inventor_country = []
    inventor_nationality = []

    applicant_name = []
    applicant_address = []
    applicant_country = []
    applicant_nationality = []

    # Data from Application Status link
    Application_type = []
    Email = []
    Additional_email = []
    Email_updated = []
    Request_for_examination_date = []
    Application_Status = []

    r = driver.find_elements(By.XPATH, '//*[@id="tableData"]/tbody/tr')
    c = driver.find_elements(By.XPATH, '//*[@id="tableData"]/tbody/tr[1]/td')

    rc = len(r)
    cc = len(c)

    for i in range(1, rc + 1, 1):
        for j in range(1, cc + 1, 1):
            d = driver.find_element(
                By.XPATH, "//tr["+str(i)+"]/td["+str(j)+"]")
            if (j == 1):
                print(d.text)
                Application_Number.append(d.text)

                # Clicking on Application Number link and switching to that window.
                d.click()
                # print(driver.current_url)
                if len(driver.window_handles) < 2:
                    print(
                        "There are not enough window handles to switch to the second window.")
                    publication_number2.append("NA")
                    publication_date2.append("NA")
                    publication_type2.append("NA")
                    application_filing_date2.append("NA")
                    priority_number2.append("NA")
                    priority_country2.append("NA")
                    priority_date2.append("NA")
                    field_of_invention2.append("NA")
                    classification_ipc2.append("NA")
                    inventor_name.append("NA")
                    inventor_address.append("NA")
                    inventor_country.append("NA")
                    inventor_nationality.append("NA")
                    applicant_name.append("NA")
                    applicant_address.append("NA")
                    applicant_country.append("NA")
                    applicant_nationality.append("NA")
                    Application_type.append("NA")
                    Email.append("NA")
                    Additional_email.append("NA")
                    Email_updated.append("NA")
                    Request_for_examination_date.append("NA")
                    Application_Status.append("NA")
                else:
                    driver.switch_to.window(driver.window_handles[1])
                    wait = WebDriverWait(driver, 30)
                    element = wait.until(EC.presence_of_element_located(
                        (By.CLASS_NAME, 'tab-content')))
                    # print(driver.current_url)

                    row_ele = driver.find_elements(
                        By.XPATH, '//*[@id="home"]/table/tbody/tr')
                    col_ele = driver.find_elements(
                        By.XPATH, '//*[@id="home"]/table/tbody/tr[1]/td')
                    rows = len(row_ele)
                    # cols = len(col_ele)

                    for k in range(1, 12, 1):
                        ele = driver.find_element(
                            By.XPATH, "//*[@id='home']/table/tbody/tr["+str(k)+"]/td[2]")
                        if (k == 2):
                            publication_number2.append(ele.text)
                        if (k == 3):
                            publication_date2.append(ele.text)
                        if (k == 4):
                            publication_type2.append(ele.text)
                        if (k == 6):
                            application_filing_date2.append(ele.text)
                        if (k == 7):
                            priority_number2.append(ele.text)
                        if (k == 8):
                            priority_country2.append(ele.text)
                        if (k == 9):
                            priority_date2.append(ele.text)
                        if (k == 10):
                            field_of_invention2.append(ele.text)
                        if (k == 11):
                            classification_ipc2.append(ele.text)

                    row_ele_inventor = driver.find_elements(
                        By.XPATH, '//*[@id="home"]/table/tbody/tr[13]/td/table/tbody/tr')
                    rows_inventor = len(row_ele_inventor)
                    # print("No. of rows in inventor is: ", rows_inventor)
                    cols_inventor = 4
                    if (rows_inventor != 0):
                        col_ele_inventor = driver.find_elements(
                            By.XPATH, '//*[@id="home"]/table/tbody/tr[13]/td/table/tbody/tr[2]/td')
                        cols_inventor = len(col_ele_inventor)
                    # print("No. of cols in inventor is: ", cols_inventor)
                    str_invname = ""
                    str_invaddress = ""
                    str_invcountry = ""
                    str_invnationality = ""

                    for m in range(2, rows_inventor + 1, 1):
                        for n in range(1, cols_inventor + 1, 1):
                            element = driver.find_element(
                                By.XPATH, "//*[@id='home']/table/tbody/tr[13]/td/table/tbody/tr["+str(m)+"]/td["+str(n)+"]")
                            if (n == 1):
                                if (m == rows_inventor):
                                    str_invname = (
                                        str_invname + str(m - 1) + ". " + element.text)
                                else:
                                    str_invname = (
                                        str_invname + str(m - 1) + ". " + element.text + " " + "\n")
                            if (n == 2):
                                if (m == rows_inventor):
                                    str_invaddress = (
                                        str_invaddress + str(m - 1) + ". " + element.text)
                                else:
                                    str_invaddress = (
                                        str_invaddress + str(m - 1) + ". " + element.text + " " + "\n")
                            if (n == 3):
                                if (m == rows_inventor):
                                    str_invcountry = (
                                        str_invcountry + str(m - 1) + ". " + element.text)
                                else:
                                    str_invcountry = (
                                        str_invcountry + str(m - 1) + ". " + element.text + " " + "\n")
                            if (n == 4):
                                if (m == rows_inventor):
                                    str_invnationality = (
                                        str_invnationality + str(m - 1) + ". " + element.text)
                                else:
                                    str_invnationality = (
                                        str_invnationality + str(m - 1) + ". " + element.text + " " + "\n")

                    inventor_name.append(str_invname)
                    inventor_address.append(str_invaddress)
                    inventor_country.append(str_invcountry)
                    inventor_nationality.append(str_invnationality)

                    row_ele_applicant = driver.find_elements(
                        By.XPATH, '//*[@id="home"]/table/tbody/tr[15]/td/table/tbody/tr')
                    rows_applicant = len(row_ele_applicant)
                    # print("No. of rows in applicant is: ", rows_applicant)
                    cols_applicant = 4
                    if (rows_applicant != 0):
                        col_ele_applicant = driver.find_elements(
                            By.XPATH, '//*[@id="home"]/table/tbody/tr[15]/td/table/tbody/tr[2]/td')
                        cols_applicant = len(col_ele_applicant)
                    # print("No. of cols in applicant is: ", cols_applicant)
                    str_appname = ""
                    str_appaddress = ""
                    str_appcountry = ""
                    str_appnationality = ""

                    for m in range(2, rows_applicant + 1, 1):
                        for n in range(1, cols_applicant + 1, 1):
                            element = driver.find_element(
                                By.XPATH, "//*[@id='home']/table/tbody/tr[15]/td/table/tbody/tr["+str(m)+"]/td["+str(n)+"]")
                            if (n == 1):
                                if (m == rows_applicant):
                                    str_appname = (
                                        str_appname + str(m - 1) + ". " + element.text)
                                else:
                                    str_appname = (
                                        str_appname + str(m - 1) + ". " + element.text + " " + "\n")
                            if (n == 2):
                                if (m == rows_applicant):
                                    str_appaddress = (
                                        str_appaddress + str(m - 1) + ". " + element.text)
                                else:
                                    str_appaddress = (
                                        str_appaddress + str(m - 1) + ". " + element.text + " " + "\n")
                            if (n == 3):
                                if (m == rows_applicant):
                                    str_appcountry = (
                                        str_appcountry + str(m - 1) + ". " + element.text)
                                else:
                                    str_appcountry = (
                                        str_appcountry + str(m - 1) + ". " + element.text + " " + "\n")
                            if (n == 4):
                                if (m == rows_applicant):
                                    str_appnationality = (
                                        str_appnationality + str(m - 1) + ". " + element.text)
                                else:
                                    str_appnationality = (
                                        str_appnationality + str(m - 1) + ". " + element.text + " " + "\n")
                    applicant_name.append(str_appname)
                    applicant_address.append(str_appaddress)
                    applicant_country.append(str_appcountry)
                    applicant_nationality.append(str_appnationality)

                    # Clicking on Application Status link inside Application Number webpage and switching to that window.
                    l = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="home"]/table/tbody/tr[18]/td/input'))).find_element(
                        By.XPATH, '//*[@id="home"]/table/tbody/tr[18]/td/input')
                    l.click()
                    driver.switch_to.window(driver.window_handles[2])
                    wait = WebDriverWait(driver, 30)
                    element = wait.until(EC.presence_of_element_located(
                        (By.CLASS_NAME, 'table-striped')))

                    row_ele = driver.find_elements(
                        By.XPATH, '//*[@id="Content"]/div[2]/table/tbody/tr')
                    col_ele = driver.find_elements(
                        By.XPATH, '//*[@id="Content"]/div[2]/table/tbody/tr[2]/td')
                    rows = len(row_ele)
                    cols = len(col_ele)
                    # print("No. of rows in application status is: ", rows)
                    # print("No. of cols in application status is: ", cols)

                    for k in range(1, rows + 1, 1):
                        ele = driver.find_element(
                            By.XPATH, "//tr["+str(k)+"]/td[2]")
                        if (k == 3):
                            # print(ele.text)
                            Application_type.append(ele.text)
                        elif (k == 8):
                            # print(ele.text)
                            Email.append(ele.text)
                        elif (k == 9):
                            # print(ele.text)
                            Additional_email.append(ele.text)
                        elif (k == 10):
                            # print(ele.text)
                            Email_updated.append(ele.text)
                        elif (k == 12):
                            # print(ele.text)
                            Request_for_examination_date.append(ele.text)

                    status_element = driver.find_element(
                        By.XPATH, '//*[@id="Content"]/div[3]/table/tbody/tr[2]/td[2]')
                    Application_Status.append(status_element.text)

                    driver.close()
                    driver.switch_to.window(driver.window_handles[1])
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])

            elif (j == 2):
                # print(d.text)
                Title.append(d.text)
            elif (j == 3):
                # print(d.text)
                Application_Date.append(d.text)
            elif (j == 4):
                # print(d.text)
                Status.append(d.text)

        # break

    # print("Error")
    # print(len(Application_Number))
    # print(len(Title))
    # print(len(Application_Date))
    # print(len(Status))
    df_t = pd.DataFrame(list(zip(Application_Number, Title, Application_Date, Status, publication_number2, publication_date2, publication_type2, application_filing_date2, priority_number2, priority_country2, priority_date2, field_of_invention2, classification_ipc2, inventor_name, inventor_address, inventor_country, inventor_nationality, applicant_name, applicant_address, applicant_country, applicant_nationality, Application_type, Email, Additional_email, Email_updated, Request_for_examination_date, Application_Status)), columns=[
                        'Application Number', 'Title', 'Application Date', 'Status', 'Publication Number', 'Publication Date(U/S 11A)', 'Publication Type', 'Application Filing Date', 'Priority Number', 'Priority Country', 'Priority Date', 'Field Of Invention', 'Classification (IPC)', 'Inventor Name', 'Inventor Address', 'Inventor Country', 'Inventor Nationality', 'Applicant Name', 'Applicant Address', 'Applicant Country', 'Applicant Nationality', 'Application Type', 'E-MAIL (As Per Record)', 'ADDITIONAL-EMAIL (As Per Record)', 'E-MAIL (UPDATED Online)', 'REQUEST FOR EXAMINATION DATE', 'Application Status'])
    master_df = pd.concat([master_df, df_t], ignore_index=True)
    out_path = "D:\\web_scraping_project\\output.xlsx"
    writer = pd.ExcelWriter(out_path, engine='xlsxwriter')
    master_df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.save()
    master_df.to_csv("D:\web_scraping_project\output_csv.csv", index=False)
    l = driver.find_element(
        By.XPATH, '//*[@id="header"]/div[4]/div/div[1]/form/div/button[3]')
    l.click()

    # break


driver.quit()
print(master_df)
print(f"Length of the dataframe is : {(master_df.shape[0])}")


# out_path = "D:\\web_scraping_project\\output.xlsx"
# writer = pd.ExcelWriter(out_path, engine='xlsxwriter')
# master_df.to_excel(writer, index=False, sheet_name='Sheet1')
# writer.save()
# master_df.to_csv("D:\web_scraping_project\output_csv.csv", index=False)
