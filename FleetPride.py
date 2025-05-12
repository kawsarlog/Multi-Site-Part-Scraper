import csv
import pandas as pd
import re
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import traceback 
from seleniumbase import Driver
import os
os.makedirs('output', exist_ok=True)

xlsx_file = "Parts for Xref 4.22.25.xlsx"  # Replace with your actual file name

driver = None
output_file_name = 'output/fleetpride_output.csv'

def normalize(text):# Lowercase, remove special characters and spaces
    if not text:
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', text.lower().strip())

def clean_and_lower(text):
    # Keep letters, numbers, and spaces — remove special characters
    cleaned = re.sub(r'[^a-zA-Z0-9 ]', '', text)  # note the space after 9
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()  # normalize multiple spaces
    return cleaned.lower()

def word_exists(a: str, b: str) -> bool:
    a = clean_and_lower(a)
    b = clean_and_lower(b)
    a_words = set(a.lower().split())
    b_words = set(b.lower().split())
    return not a_words.isdisjoint(b_words)

def save_data(data):
    while len(data) < 18:
        data.append("N/A")
    data = data[:18]  # Ensure data has exactly 18 elements
    with open(output_file_name, mode='a', newline='',encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow([
        "Part Number Reformat", "DESCRIPTION", "Invoice Reference",
        "fleetpride Part # 1", "fleetpride Desc 1", "fleetpride Price 1", "fleetpride CS 1", "fleetpride Photo 1",
        "fleetpride Part # 2", "fleetpride Desc 2", "fleetpride Price 2", "fleetpride CS 2", "fleetpride Photo 2",
        "fleetpride Part # 3", "fleetpride Desc 3", "fleetpride Price 3", "fleetpride CS 3", "fleetpride Photo 3"
    ])
        writer.writerow(data)

def visibil_element(driver, by, selector, wait=10): ### web element find and search
    element = False
    if by == 'xpath':
        byselector = By.XPATH
    try:
        element = WebDriverWait(driver, wait).until(
            EC.visibility_of_element_located((byselector, selector)))
    except Exception as e:
        element = False
    return element

def driver_define():
    try:
        driver = Driver(uc=True)
        driver.maximize_window()
        time.sleep(2)
        return driver
    except Exception as e:
        print(traceback.format_exc())

def driver_opening():
    global driver
    faild_opening = 0
    while True:
        try:
            faild_opening += 1
            print("Trying opening driver >>  ", faild_opening)
            try:driver.quit()
            except:pass
            driver_open = False
            driver = driver_define()
            
            for tr in range(6):
                print("Trying getting website : " , tr + 1 )
                try:
                    driver.get('https://www.fleetpride.com/gs/%40uri#q=otr105771')
                    driver.get('https://www.fleetpride.com/gs/%40uri#q=otr105771')
                except:
                    try:
                        driver.get('https://www.fleetpride.com/gs/%40uri#q=otr105771')
                        driver.get('https://www.fleetpride.com/gs/%40uri#q=otr105771')
                    except:break  
                time.sleep(3)
                try:error_ck = driver.find_element(By.XPATH , '//div[@class="error-title"]')
                except:error_ck = False
                if error_ck:
                    print("Error page")
                    driver.quit()
                    break
                print(driver.title)
                if 'Search' in driver.title:
                    print("Browser open success")
                    driver_open = True
                    break
            if driver_open:
                break
            try:driver.quit()
            except:pass
        except Exception as e:
            print(traceback.format_exc())

def gatting_page(search_url):
    global driver
    while True:
        try:
            try:
                driver.get(search_url)
                driver.get(search_url)
            except:
                try:
                    driver.get(search_url)
                    driver.get(search_url)
                except:break
            time.sleep(2)
            if 'Search' in driver.title:
                return True
            try:
                time.sleep(2)
                driver.find_element(By.XPATH , '//div[@class="error-title"]')
                pass
            except:
                return True
            print("Page getting Failed..")
            driver.quit()
            driver_opening()
        except:
            print(traceback.format_exc())
            driver.quit()
            driver_opening()

def get_list_data(search_number,search_disc,e_col,d_col): # data is list = [sku,title,pprice,img]
    
    # Normalize & clean
    search_number_n, search_number_c = normalize(search_number), clean_and_lower(search_number)
    d_col_c,d_col_n = clean_and_lower(d_col), normalize(d_col)
    e_col_n, e_col_c = normalize(e_col), clean_and_lower(e_col)

    search_number_n = normalize(search_number)
    global driver
    all_products = []
    match_count = 0
    res_f = False
    url = f"https://www.fleetpride.com/gs/%2540uri#q={search_number_n}"
    for _ in range(3):
        try:
            gatting_page(url)
            time.sleep(3) #//div[@class="resultCount"] //span[@part="formatted-rich-text"]
            
            for _ in range(15):
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH , '//div[@class="resultCount"]')
                    res_f = True
                    print("Result found...")
                    time.sleep(3)
                    try:
                        driver.find_element(By.XPATH , '//button[@data-view="list"]').click()
                        time.sleep(2)
                    except:pass
                    break
                except:
                    try:
                        ch = driver.find_elements(By.XPATH , '//span[@part="formatted-rich-text"]')
                        if len(ch) == 2:
                            print("No Result found...")
                            return []
                    except:
                        pass
            if res_f:
                
                total_res = driver.find_elements(By.XPATH , '//div[@class="slds-grid plp-card plp slds-grid_vertical-stretch"]')
                
                print("total result : ",len(total_res))
                for xxxx in range(1,len(total_res)+1):
                    if match_count == 3:
                        break
                    try:
                        matchh = driver.find_element(By.XPATH , f'(//div[@class="slds-grid plp-card plp slds-grid_vertical-stretch"])[{xxxx}]//span[@class="cross-reference-match"]').text
                        if 'Part # Match' in matchh:
                            print("Exact match found")
                            
                            sku = driver.find_element(By.XPATH , f'(//div[@class="slds-grid plp-card plp slds-grid_vertical-stretch"])[{xxxx}]//span[@class="result-text__value"]').text
                            try:title = driver.find_element(By.XPATH , f'(//div[@class="slds-grid plp-card plp slds-grid_vertical-stretch"])[{xxxx}]//span[@class="highlighted-title"]').text
                            except:title = ''
                            try:price = driver.find_element(By.XPATH , f'(//div[@class="slds-grid plp-card plp slds-grid_vertical-stretch"])[{xxxx}]//div[@class="your-price"]').text
                            except:price = ''
                            try:img = driver.find_element(By.XPATH , f'(//div[@class="slds-grid plp-card plp slds-grid_vertical-stretch"])[{xxxx}]//img[@loading="lazy"]').get_attribute('src')
                            except:img = ''
                            word_exi = word_exists(search_disc,title)
                            if word_exi:
                                all_products.append(sku)
                                all_products.append(title)
                                all_products.append(price)
                                all_products.append('Exact match')
                                all_products.append(img)
                                match_count += 1
                            continue
                        
                        if 'Partial' in matchh or 'X-Ref' in matchh:
                            print("X-Ref match found")
                            sku = driver.find_element(By.XPATH , f'(//div[@class="slds-grid plp-card plp slds-grid_vertical-stretch"])[{xxxx}]//span[@class="result-text__value"]').text
                            try:title = driver.find_element(By.XPATH , f'(//div[@class="slds-grid plp-card plp slds-grid_vertical-stretch"])[{xxxx}]//span[@class="highlighted-title"]').text
                            except:title = ''
                            try:price = driver.find_element(By.XPATH , f'(//div[@class="slds-grid plp-card plp slds-grid_vertical-stretch"])[{xxxx}]//div[@class="your-price"]').text
                            except:price = ''
                            try:img = driver.find_element(By.XPATH , f'(//div[@class="slds-grid plp-card plp slds-grid_vertical-stretch"])[{xxxx}]//img[@loading="lazy"]').get_attribute('src')
                            except:img = ''
                            word_exi = word_exists(search_disc,title)
                            if word_exi:
                                all_products.append(sku)
                                all_products.append(title)
                                all_products.append(price)
                                all_products.append('X-ref match')
                                all_products.append(img)
                                match_count += 1
                            continue
                        
                        try:sku = driver.find_element(By.XPATH , f'(//div[@class="slds-grid plp-card plp slds-grid_vertical-stretch"])[{xxxx}]//span[@class="result-text__value"]').text
                        except:continue
                        try:title = driver.find_element(By.XPATH , f'(//div[@class="slds-grid plp-card plp slds-grid_vertical-stretch"])[{xxxx}]//span[@class="highlighted-title"]').text
                        except:title = ''
                        try:price = driver.find_element(By.XPATH , f'(//div[@class="slds-grid plp-card plp slds-grid_vertical-stretch"])[{xxxx}]//div[@class="your-price"]').text
                        except:price = ''
                        try:img = driver.find_element(By.XPATH , f'(//div[@class="slds-grid plp-card plp slds-grid_vertical-stretch"])[{xxxx}]//img[@loading="lazy"]').get_attribute('src')
                        except:img = ''
                        word_exi = word_exists(search_disc,title)
                        if word_exi:
                            sku_n, sku_c = normalize(sku), clean_and_lower(sku)
                            exact_conditions = [
                                    e_col_n == sku_n, 
                                    e_col_c == sku_c,
                                    d_col_n == sku_n, 
                                    d_col_c == sku_c,
                                    search_number_n == sku_n, 
                                    search_number_c == sku_c
                                ]
                            if any(exact_conditions):
                                all_products.append(sku)
                                all_products.append(title)
                                all_products.append(price)
                                all_products.append('Exact match')
                                all_products.append(img)
                                match_count += 1
                                continue
                            partial_conditions = [
                                    e_col_n in sku_n, 
                                    e_col_c in sku_c,
                                    d_col_n in sku_n, 
                                    d_col_c in sku_c,
                                    search_number_n in sku_n, 
                                    search_number_c in sku_c
                                ]
                            if any(partial_conditions) and not search_number_c.isdigit():
                                all_products.append(sku)
                                all_products.append(title)
                                all_products.append(price)
                                all_products.append('Partial match')
                                all_products.append(img)
                                match_count += 1
                                continue
                                
                    except:
                        pass
            else:
                print("No Result found...")
                return []
            return all_products
        except Exception as e:
            # print(e)
            driver.quit()
            driver_opening()
            
    return []

def collect_details(search_number, search_disc,e_col,d_col):
    global driver
    
    data = get_list_data(search_number,search_disc,e_col,d_col)
    if data is False:
        with open("m_error.txt", 'a') as f:
            f.write(f"{search_number} - {search_disc}\n")
        return False

    final_data = [search_number, search_disc, 'N/A']

    for m in data:
        final_data.append(m)  # SKU

    save_data(final_data)

df = pd.read_excel(xlsx_file, engine="openpyxl")  # Read first sheet

df_filtered = df[0:]

driver_opening()

for index, row in df_filtered.iterrows():
    print(f">>>>> : {index + 1}")
    search_number = row.tolist()[1]
    search_number = str(search_number).strip() if search_number is not None else None

    search_disc = row.tolist()[2]
    search_disc = str(search_disc).strip() if search_disc is not None else None

    d_col = row.tolist()[4]
    d_col = str(d_col).strip() if d_col is not None else None

    e_col = row.tolist()[5]
    e_col = str(e_col).strip() if e_col is not None else None

    print(">> ",search_number," >>> ", search_disc)
    
    collect_details(search_number, search_disc,e_col,d_col)
    
print("Script Completed...")