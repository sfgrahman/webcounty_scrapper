from selenium import webdriver
from chromedriver_py import binary_path 
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from time import sleep
import pandas as pd
from bs4 import BeautifulSoup
import re
from random import randint
from datetime import datetime
from time import perf_counter
# from openpyxl import Workbook


t0 = perf_counter()
current_datetime = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
raw_data = []
processed_data = []
date = "04/25/2024" #mm/dd/yyyy
bookv = "Official Public Records"
raw_column_header = ['BASEM', 'Reception #', 'BOOK_VOL', 'Book_Page', 'Doc Type', 'Recorded Date', 'GRANTOR (Name)', 'GRANTEE (Other Name)']
processed_column_header = ['TOWNSHIP', 'SECTION', 'Reception #', 'BOOK_VOL', 'Book', 'Page', 'Doc Type', 'Recorded Date', 'GRANTOR (Name)', 'GRANTEE (Other Name)']

# Define a custom user agent
my_user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
"""
# openpyxl implementation code
file = f'excel_{str(current_datetime)}.xlsx'
Load existing workbook
wb = Workbook()
Select the active sheet
ws = wb.active
ws.append(column_header)
"""

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ['enable-automation'])
options.add_argument(f"--user-agent={my_user_agent}")

svc = webdriver.ChromeService(executable_path=binary_path)
driver = webdriver.Chrome(service=svc, options=options)

driver.get("https://countyfusion2.govos.com/")

# click on the countylink
for i in range(10):
    try:
        countylink = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable(driver.find_element(By.XPATH, "//*[contains(text(), 'Eddy County Clerk')]"))
            )
        driver.execute_script("arguments[0].click();", countylink)
        sleep(randint(1, 3))
        break
    except NoSuchElementException as e:
        sleep(randint(1, 3))
        driver.refresh()
else:
    raise e

# click on the guestlogin
for i in range(10):
    try:
        guestlogin = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable(driver.find_element(By.XPATH, "//input[@value='Login as Guest']"))
            )
        driver.execute_script("arguments[0].click();", guestlogin)
        sleep(randint(1, 3))
        break
    except NoSuchElementException as e:
        sleep(randint(1, 3))
        driver.refresh()
else:
    raise e

# click on the publicrecord
for i in range(10):
    try:
        publicrecord = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable(driver.find_element(By.XPATH, "//*[contains(text(), 'Search Public Records')]"))
            )
        driver.execute_script("arguments[0].click();", publicrecord)
        sleep(randint(1, 3))
        break
    except NoSuchElementException as e:
        sleep(randint(1, 3))
        driver.refresh()
else:
    raise e

sleep(5) 
# enter into body frame to search dynamic content
frame = WebDriverWait(driver, 300).until(
	EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[@name='bodyframe']"))
)
sleep(2) 
# enter into dynamic search frame 
subframe = WebDriverWait(driver, 300).until(
	EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[@name='dynSearchFrame']"))
)
# with open("dynSearchFrame.html", 'w') as e:
#     e.write(driver.page_source)

for i in range(10):
    try:
        legal_description = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable(driver.find_element(By.XPATH, "//*[contains(text(), 'Legal Description')]"))
            )
        driver.execute_script("arguments[0].click();", legal_description)
        sleep(randint(1, 3))
        break
    except NoSuchElementException as e:
        sleep(randint(1, 3))
        driver.refresh()
else:
    raise e
# # click on only oil and gas
# for i in range(10):
#     try:
#         oilgas = WebDriverWait(driver, 300).until(
#                 EC.element_to_be_clickable(driver.find_element(By.XPATH, "//*[@id='nameTree_easyui_tree_8']/span[4]"))
#             )
#         driver.execute_script("arguments[0].click();", oilgas)
#         sleep(randint(1, 3))
#         break
#     except NoSuchElementException as e:
#         sleep(randint(1, 3))
#         driver.refresh()
# else:
#     raise e

sleep(2) 
# enter into another child frame to enter data
subframe2 = WebDriverWait(driver, 300).until(
	EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[@name='criteriaframe']"))
)
# Enter the data and press enter to search dynamic content
for i in range(10):
    try:
        date_field = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable(driver.find_element(By.ID, "_easyui_textbox_input10"))
            ).send_keys(date, Keys.ENTER)
        sleep(randint(1, 3))
        break
    except NoSuchElementException as e:
        sleep(randint(1, 3))
        driver.refresh()

# switch back to parent frame
driver.switch_to.default_content()
sleep(5) 
frame2 = WebDriverWait(driver, 300).until(
	EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[@name='bodyframe']"))
)

sleep(2) 
resultFrame = WebDriverWait(driver, 300).until(
	EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[@name='resultFrame']"))
)

# Find page number
pn = BeautifulSoup(driver.page_source, 'html.parser')
nav = pn.find('div', id='navLabelDisplay')
k = nav.tbody.tbody.tr.td.text
page_no = int(" ".join(k.split()).replace('Page 1 of ', ''))

# def preprocess_text(txt):
    ## remove whitespaces
#     return re.sub("\s\s+", " ", txt)

def preprocess_basesm(bas, reception, bookv, book_page, doc_type, recorded, name, other_name):
    # preprocess all docuemtn as par client requirements
    sec_no = []
    try:
        if all([x in bas for x in ['SEC', 'TSHP', 'RANGE']]):
            try:
                other_name = other_name.split('|br|')
            except:
                other_name = ''
                pass
            try:
                name = name.split('|br|')
                if len(name) > 1:
                    name = f"{name[0]} ET AL"
                if len(name) == 1:
                    name = name[0]
            except:
                name = ''
                pass
                
            bp = book_page.split()

            try:
                sec = re.search('SEC\s[\d\s,-]+', bas).group().replace('SEC ', '')
                sec = sec.replace(' ', '')
                tshp = re.search('TSHP\s[\d\s,-]+', bas).group().replace('TSHP ', '')
                tshp = tshp.replace(' ', '')
                rng = re.search('RANGE\s[\d\s,-]+', bas).group().replace('RANGE ', '')
                rang = f"{tshp}S;{rng}E"
                # LEASE NMNM 0455265 SEC 1, 9-15, 20, 21 TSHP 20 RANGE 27
                if "," in sec:
                    k = sec.split(",")
                    k = list(filter(None, k))
                    for kk in k:
                        if "-" in kk:
                            sp = kk.split('-')
                            for jj in range(int(sp[0]), int(sp[1])+1):
                                sec_no.append(jj)
                        if kk != " ":
                            if "-" not in kk:
                                sec_no.append(int(kk))            
                if "-" in sec:
                    if "," not in sec:
                        sp1 = sec.split('-')
                        for kks in range(int(sp1[0]), int(sp1[1])+1):
                            sec_no.append(kks) 
                if all([x not in sec for x in [',', '-']]):
                    sec_no.append(int(sec))
                # print(sec_no)
            except Exception as e:
                # print(bas)
                print(e)
                pass

            
            if other_name == "":
                for sn in sec_no:
                    if [rang, sn, reception, bookv, bp[0], bp[1], doc_type, recorded, name, ""] not in processed_data:
                        processed_data.append([rang, sn, reception, bookv, bp[0], bp[1], doc_type, recorded, name, ""])
            else:
                for other in other_name:
                    for sn in sec_no:
                        if [rang, sn, reception, bookv, bp[0], bp[1], doc_type, recorded, name, other] not in processed_data:
                            processed_data.append([rang, sn, reception, bookv, bp[0], bp[1], doc_type, recorded, name, other])
                
    except Exception as e:
        # print(e)
        pass
# add page_no in for loop to scrape all page
for p in range(page_no):
    sleep(2) 
    resultListFrame = WebDriverWait(driver, 300).until(
        EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[@name='resultListFrame']"))
    )
    sleep(5) 
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    for tr in soup.find_all('tr', {'id': re.compile(r'datagrid-row-r1-2-\d+')}):
        try:
            reception = tr.find('div', "datagrid-cell-c1-3").text.strip()
            book_page = tr.find('div', "datagrid-cell-c1-4").text.strip()
            # name = tr.find('div', "datagrid-cell-c1-6").text.strip()
            name = tr.find('div', "datagrid-cell-c1-6").get_text(separator='|br|', strip=True)
            # other_name = tr.find('div', "datagrid-cell-c1-8").text.strip()
            other_name = tr.find('div', "datagrid-cell-c1-8").get_text(separator='|br|', strip=True)

            doc_type = tr.find('div', "datagrid-cell-c1-9").text.strip()
            recorded = tr.find('div', "datagrid-cell-c1-10").text.strip()
            additional_data = tr.find('td', {"field": "additionalData"})
            basesm = additional_data.find_all('div', {"class": "basesm"})
            for b in basesm:
                bas = b.text.strip()
                raw_data.append([bas, reception, bookv, book_page, doc_type, recorded, name, other_name])
                preprocess_basesm(bas, reception, bookv, book_page, doc_type, recorded, name, other_name)
                # print(bas)
                # ws.append([bas, reception, bookv, book_page, doc_type, recorded, name, other_name])
            # wb.save(file)
        except Exception as e:
            print(e)
            if bas != "":
                with open(f'Exception_{str(current_datetime)}.txt', 'a') as ap:
                    ap.writelines([bas, reception, bookv, book_page, doc_type, recorded, name, other_name])
            pass
    sleep(2)
    if p != page_no -1:
        driver.switch_to.parent_frame()

        resultFrame2 = WebDriverWait(driver, 300).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[@name='subnav']"))
        )
        for i in range(10):
            try:
                alink = WebDriverWait(driver, 300).until(
                        EC.element_to_be_clickable(driver.find_element(By.XPATH, "//img[@title='Go to next result page']"))
                    )
                driver.execute_script("arguments[0].click();", alink)
                break
            except NoSuchElementException as e:
                sleep(randint(1, 3))
            # driver.refresh()
        
        driver.switch_to.parent_frame()

sleep(10) 
driver.close()

# remove dublicate

def convert_to_excel(data, column_header, filename, mode):
    df = pd.DataFrame(data, columns =column_header)
    if mode == 'Y':
        df.drop_duplicates(inplace=True)
    df.to_excel(f"{filename}_{str(current_datetime)}.xlsx", index=False)

convert_to_excel(raw_data, raw_column_header, "raw_countyfusion2", "N")
print(f"Raw data based `raw_countyfusion2_{str(current_datetime)}.xlsx` excel file created successfully.")
convert_to_excel(processed_data, processed_column_header, "processed_countyfusion2", "Y")
print(f"Processed data based `processed_countyfusion2_{str(current_datetime)}.xlsx` excel file created successfully.")
t1 = perf_counter()
print("Total time taken: ", t1 - t0)