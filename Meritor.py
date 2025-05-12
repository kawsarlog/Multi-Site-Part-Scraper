import time
import csv
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import traceback 
from seleniumbase import Driver
import pandas as pd
import re
import os
os.makedirs('output', exist_ok=True)

xlsx_file = "Parts for Xref 4.22.25.xlsx"  # Replace with your actual file name

output_file = "output/meritorpartsxpress_output.csv"
driver = None

def driver_define():
    try:
        driver = Driver(uc=True)
        driver.maximize_window()
        time.sleep(2)
        return driver
    except Exception as e:
        print(traceback.format_exc())

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
                    driver.get('https://www.meritorpartsxpress.com/webapp/wcs/stores/servlet/en/meritor-na/home')
                    driver.get('https://www.meritorpartsxpress.com/webapp/wcs/stores/servlet/en/meritor-na/home')
                except:
                    try:
                        driver.get('https://www.meritorpartsxpress.com/webapp/wcs/stores/servlet/en/meritor-na/home')
                        driver.get('https://www.meritorpartsxpress.com/webapp/wcs/stores/servlet/en/meritor-na/home')  
                    except:break  
                time.sleep(3)
                try:error_ck = driver.find_element(By.XPATH , '//div[@class="error-title"]')
                except:error_ck = False
                if error_ck:
                    print("Error page")
                    driver.quit()
                    break
                print(driver.title)
                if 'MPX Home Page - Meritor Aftermarket NA Store is your Xpress Lane for Uptime' in driver.title:
                    print("Browser open success")
                    driver_open = True
                    try:
                        visibil_element(driver, 'xpath', '//button[@id="onetrust-accept-btn-handler"]', wait=15)
                        driver.find_element(By.XPATH , '//button[@id="onetrust-accept-btn-handler"]').click()
                        time.sleep(1)
                    except:pass
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
            if 'Search Results' in driver.title:
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

def save_data(data):
    
    while len(data) < 18:
        data.append("N/A")
    with open(output_file, mode='a', newline='',encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow([
        "Part Number Reformat", "DESCRIPTION", "Invoice Reference",
        "meritorpartsxpress  Part # 1", "meritorpartsxpress  Desc 1", "meritorpartsxpress  Price 1", "meritorpartsxpress  CS 1", "meritorpartsxpress  Photo 1",
        "meritorpartsxpress  Part # 2", "meritorpartsxpress  Desc 2", "meritorpartsxpress  Price 2", "meritorpartsxpress  CS 2", "meritorpartsxpress  Photo 2",
        "meritorpartsxpress  Part # 3", "meritorpartsxpress  Desc 3", "meritorpartsxpress  Price 3", "meritorpartsxpress  CS 3", "meritorpartsxpress  Photo 3"
    ])
        writer.writerow(data)

def find_best_matches(data, search_number, search_disc, d_col, e_col):
    collect_data = []
    data_found = 0
    max_data = 3

    # Normalize & clean
    search_number_n, search_number_c = normalize(search_number), clean_and_lower(search_number)
    search_disc_c = clean_and_lower(search_disc)
    d_col_c = clean_and_lower(d_col)
    d_col_n = normalize(d_col)
    e_col_n, e_col_c = normalize(e_col), clean_and_lower(e_col)

    for item in data:
        if data_found == max_data:
            break
        if len(item) < 2:
            continue  # Invalid row

        sku = item[0].strip().lower()
        title = item[1].strip().lower()

        if not sku or not title:
            continue
        sku_n, sku_c = normalize(sku), clean_and_lower(sku)
        title_c = clean_and_lower(title)

        matched = False
        matching_res = ""
        
        # ✅ Exact Match
        exact_conditions = [
            e_col_n in [sku_n], e_col_c in [sku_c],
            d_col_n in [sku_n], d_col_c in [sku_c],
            search_number_n in [sku_n], search_number_c in [sku_c]
        ]

        if search_number_n in [sku_n]:
            print("sn in sku_n")
        if any(exact_conditions) and (word_exists(title_c, search_disc_c) or word_exists(title_c, d_col_c)):
            matching_res = 'Exact Match'
            matched = True
            print(f"[Exact] Matched SKU: {sku} | Title: {title}")

        # ✅ x-ref match
        elif any([
            word_exists(e_col_c, sku_c),
            word_exists(d_col_c, sku_c),
            word_exists(search_number_c, sku_c)
        ]) and (word_exists(title_c, search_disc_c) or word_exists(title_c, d_col_c)):
            matching_res = 'x-ref match'
            matched = True
            print(f"[x-ref] Substring Match with SKU: {sku} | Title: {title}")

        if matched:
            collect_data.extend([
                [
                    item[0],
                    item[1],
                    item[2],
                    matching_res,
                    item[3]
                ]
            ])
            data_found += 1

    print("Collected Matches:", collect_data)
    return collect_data

def get_list_data(search_number,start,page_number): # data is list = [sku,title,pprice,img]
    global driver
    if page_number == 'search':
        page_number = 1
    if search_number.startswith('#'):
        search_number = search_number[1:]
    productBeginIndex = start
    
    url = f"https://www.meritorpartsxpress.com/webapp/wcs/stores/servlet/SearchDisplay?categoryId=63501&storeId=10154&catalogId=10051&langId=-1&sType=SimpleSearch&resultCatEntryType=2&showResultsPage=true&searchSource=Q&pageView=&beginIndex=0&pageSize=12&isRequestComingFromPrimarySearchBar=isRequestComingFromPrimarySearchBar&searchTermScope=102&searchTerm={search_number}#facet:&productBeginIndex:{productBeginIndex}&facetLimit:&orderBy:&pageView:grid&minPrice:&maxPrice:&pageSize:100&"
    for _ in range(3):
        try:
            gatting_page(url)
            time.sleep(3)
            try:
                driver.find_element(By.XPATH , '//h2[@id="itemName"]')
                return None
            except:pass
            visibil_element(driver, 'xpath', "//span[contains(@id, 'searchTotalCount')]", wait=10)
            total_items_str = driver.find_element(By.XPATH , "//span[contains(@id, 'searchTotalCount')]").text.replace('matches.', '').replace(",", '').strip()
            total_items = int(total_items_str)
            total_pages = (total_items + 100 - 1) // 100  # Calculate total pages
            print("Total Items : ", total_items)
            if total_items == 0:
                return None
            
            for _ in range(60):  # Wait up to 120 seconds (60 * 2s)
                time.sleep(2)

                
                this_page_products = driver.find_elements(By.XPATH, '//div[@class="product_info"]')
                num_products = len(this_page_products)
                print("Found  " , num_products)
                print("page_number " , page_number)
                print("total_pages " , total_pages)
                # Last page condition (should load all remaining products)
                if page_number == total_pages:
                    if num_products == total_items % 100 or num_products == total_items:
                        break  # All remaining products loaded
                    # print(f"Waiting for last page products to load... ({num_products}/{total_items % 100})")
                    continue

                # If total_items <= 100, wait for all products to load
                if total_items <= 100:
                    if num_products == total_items:
                        break
                    # print(f"Waiting for all {total_items} products to load... ({num_products}/{total_items})")
                    continue

                # If total_items > 100, wait until at least 100 products load per page
                if num_products == 100:
                    break  # Full page loaded
                # print(f"Waiting for 100 products to load... ({num_products}/100)")

            # Now proceed to the next page or process the products
            print(f"Products loaded on page {page_number}: {num_products}")
            ##
            print("Product Items : ", len(this_page_products))
            all_products = []
            for itm in range(1, len(this_page_products) + 1):
                try:number = driver.find_element(By.XPATH , f'((//div[@class="product_info"])[{itm}]//div[@class="rir-data-item name"])[1]').text
                except: number = 'N/A'
                try:disc = driver.find_element(By.XPATH , f'((//div[@class="product_info"])[{itm}]//div[@class="rir-data-item name"])[2]').text
                except: disc = 'N/A'
                try:img = driver.find_element(By.XPATH , f'(//div[@class="product_info"])[{itm}]/preceding-sibling::div[1]//img').get_attribute('src')
                except: img = 'N/A'
                all_products.append([number,disc,'N/A',img])
    
            return all_products
        except Exception as e:
            # print(e)
            driver.quit()
            driver_opening()
            
    return []

def collect_details(search_number, search_disc, d_col, e_col):
    global driver
    # Get initial data
    start = 0
    data = get_list_data(search_number, start, page_number='search')
    if data is False:
        with open("m_error.txt", 'a') as f:
            f.write(f"{search_number} - {search_disc}\n")
        return False

    try:
        visibil_element(driver, 'xpath', "//span[contains(@id, 'searchTotalCount')]", wait=10)
        total_items_str = driver.find_element(By.XPATH, "//span[contains(@id, 'searchTotalCount')]").text.replace('matches.', '').replace(",", '').strip()
        total_item = int(total_items_str)
    except:
        with open("m_error.txt", 'a') as f:
            f.write(f"{search_number} - {search_disc}\n")
        return False

    total_pages = (total_item + 100 - 1) // 100  # Calculate total pages
    print("Total Pages:", total_pages)
    best_matches = []

    # Iterate through pages until we get 3 best matches
    for page in range(total_pages):
        if len(best_matches) >= 3:
            break  # Stop if we already have 3 matches

        start = page * 100
        if page != 0:
            page_number = page + 1
            data = get_list_data(search_number, start, page_number)
            if data is False or data == []:
                with open("m_error.txt", 'a') as f:
                    f.write(f"{search_number} - {search_disc}\n")
                return False

        matches = find_best_matches(data, search_number, search_disc, d_col, e_col)
        best_matches.extend(matches)

        # Keep only the top 3 matches
        best_matches = best_matches[:3]
        print(f"Page {page + 1} - Best Matches: {len(best_matches)}")
        if len(best_matches) == 3:
            break

    final_data = [search_number, search_disc, 'N/A']

    for m in best_matches:
        for n in m:
            final_data.append(n)  # SKU

    save_data(final_data)

def extract_part_numbers():
    part_numbers = []
    try:
        with open(output_file, mode='r', newline='', encoding='utf-8') as file:
            reader = csv.reader(file)
            header = next(reader)  # Read the header row
            
            # Find the index of "Part Number Reformat"
            part_number_index = header.index("Part Number Reformat")
            
            for row in reader:
                part_numbers.append(row[part_number_index])
    except:
        part_numbers = []
    return part_numbers

driver_opening()

part_numbers = extract_part_numbers()

df = pd.read_excel(xlsx_file, engine="openpyxl")  # Read first sheet
df_filtered = df[0:]

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
    
    print(search_number," >>> ", search_disc)
    if search_number in part_numbers:
        print(f"Skipping {search_number} - {search_disc}")
        continue
    
    collect_details(search_number,search_disc,d_col,e_col)

print("Script Completed...")
