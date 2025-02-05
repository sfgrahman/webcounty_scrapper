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
from random import randint
from datetime import datetime
from time import perf_counter
from tqdm import tqdm
from openpyxl import Workbook

t0 = perf_counter()
current_datetime = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
processed_data = []
date = "04252024"   #MMDDYYYY
process_header = ['TOWNSHIP', 'SECTION', 'Others', 'TYPE', 'FILED_DATE', 'GRANTOR', 'GRANTEE', 'Instrument', 'Doc#']
base_url = "https://liveweb.leacounty-nm.org/"

grantor = ''
grantee = ''
sec = ''
tship = ''
rng = ''
other = ''
rang = ''

urls = []
err_urls_header = ['Error']
err_urls = []
total_url = ['Total Url #']
curr_url_header = ['Current URL']
curr_url = []

# openpyxl implementation code
file = f'raw_data_{str(current_datetime)}.xlsx'
# Load existing workbook
wb = Workbook()
# Select the active sheet
ws = wb.active
ws.append(process_header)

def convert_to_excel(data, column_header, filename, mode):
    df = pd.DataFrame(data, columns =column_header)
    if mode == 'Y':
        df.drop_duplicates(inplace=True)
    df.to_excel(f"{filename}_{str(current_datetime)}.xlsx", index=False)
# Define a custom user agent
my_user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ['enable-automation'])
options.add_argument(f"--user-agent={my_user_agent}")

svc = webdriver.ChromeService(executable_path=binary_path)
driver = webdriver.Chrome(service=svc, options=options)

driver.get("http://liveweb.leacounty-nm.org/Menux.aspx?source=Main")

# click on the Clerk
for i in range(10):
    try:
        clerk = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable(driver.find_element(By.XPATH, "//*[contains(text(), 'Clerk')]"))
            )
        driver.execute_script("arguments[0].click();", clerk)
        sleep(randint(1, 3))
        break
    except NoSuchElementException as e:
        sleep(randint(1, 3))
        driver.refresh()
else:
    raise e

# click on the Grantee
for i in range(10):
    try:
        grantor = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable(driver.find_element(By.XPATH, "//*[contains(text(), 'Grantor')]"))
            )
        driver.execute_script("arguments[0].click();", grantor)
        sleep(randint(1, 3))
        break
    except NoSuchElementException as e:
        sleep(randint(1, 3))
        driver.refresh()
else:
    raise e

sleep(2)
# Enter the data and press enter to search dynamic content
for i in range(10):
    try:
        date_field = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable(driver.find_element(By.NAME, "filedte"))
            ).send_keys(date, Keys.ENTER)
        sleep(randint(1, 3))
        break
    except NoSuchElementException as e:
        sleep(randint(1, 3))
        driver.refresh()
else:
    raise e

while True:
    try:
        curr_url.append(driver.current_url)
        page_content = BeautifulSoup(driver.page_source, 'html.parser')

        table = page_content.find('table', {"id": "tableResults"})
        body = table.find('tbody')
        trb = body.find_all('tr')
        for tr in trb:
            for a in tr.find_all('a', href=True):
                link = f"{base_url}{a['href']}"
                if link not in urls:
                    urls.append(link)
        # click on the Next
        for i in range(10):
            try:
                next_button = WebDriverWait(driver, 300).until(
                        EC.element_to_be_clickable(driver.find_element(By.XPATH, "//*[contains(text(), 'Next')]"))
                    )
                driver.execute_script("arguments[0].click();", next_button)
                sleep(randint(1, 3))
                break
            except NoSuchElementException as e:
                sleep(randint(1, 3))
                driver.refresh()
        else:
            raise e

    except:
        break
# convert_to_excel(curr_url, curr_url_header, 'current_url_', "Y")
# print(f"`current_url_{str(current_datetime)}` excel file created successfully.")

# convert_to_excel(urls, total_url, 'total_url_', "Y")
# print(f"`total_url_{str(current_datetime)}` excel file created successfully.")  
"""
with open("done.txt", "w") as done:
"""
for url in tqdm(urls):
    # done.write(url+'\n')
    sleep(randint(3, 7))
    driver.get(url)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    try:
        reception_no = soup.find("b", string="Reception #").next_sibling.get_text().strip()
        description = soup.find("b", string="Kind of Instrument").next_sibling.get_text().strip()
        f_date= soup.find("b", string="Date Filed").next_sibling.get_text().strip()
        filed_date = datetime.strptime(f_date, '%Y%m%d').strftime('%m/%d/%Y')
        i_date = soup.find("b", string="Intrument Date").next_sibling.get_text().strip()
        instrument = datetime.strptime(i_date, '%Y%m%d').strftime('%m/%d/%Y')

        grantee = soup.find_all("fieldset")[1].get_text(separator='|br|', strip=True)
        grtee = grantee.replace("Grantee Information", "").split('|br|')
        grtee_update = list(filter(None, grtee))
        if len(grtee_update) > 1:
            grantee = f"{grtee_update[0]} ET AL"
        else:
            grantee = grtee_update[0]

        grantor = soup.find_all("fieldset")[2].get_text(separator='|br|', strip=True)
        grtor = grantor.replace("Grantor Information", "").split('|br|')
        grtor_update = list(filter(None, grtor))
        if len(grtor_update) > 1:
            grantor = f"{grtor_update[0]} ET AL"
        else:
            grantor = grtor_update[0]

        sections = soup.find_all("b", string="Section")
        townships = soup.find_all("b", string="Township")
        ranges = soup.find_all("b", string="Range")

        for (section, township, rnge) in zip(sections, townships, ranges):
            if section != "" and township != "" and rnge != "":
                try:
                    sec = section.next_sibling.get_text().strip()
                except:
                    sec = None
                try:
                    tship = township.next_sibling.get_text().strip()
                except:
                    tship = None
                try:
                    rang = rnge.next_sibling.get_text().replace('\xa0', '').split(" ")
                    rng = list(filter(None, rang))
                except:
                    rng = None
                if rng != None:
                    rang = f"{tship};{rng[0]}"
                else:
                    rang = f"{tship};{None}"
                try:
                    other = rng[1]
                except:
                    None
                processed_data.append([rang, sec, other, description, filed_date, grantor, grantee, instrument, reception_no])
                ws.append([rang, sec, other, description, filed_date, grantor, grantee, instrument, reception_no])
            wb.save(file)
    except:
        err_urls.append(url)
        # print("Error: ", url)
        pass


convert_to_excel(processed_data, process_header, 'processed_leacounty', "Y")
print(f"Raw data based `processed_leacounty_{str(current_datetime)}` excel file created successfully.")

# convert_to_excel(err_urls, err_urls_header, 'error_urls_leacounty', "Y")
# print(f"Error `error_urls_leacounty_{str(current_datetime)}` excel file created successfully.")

sleep(10) 
driver.close()
