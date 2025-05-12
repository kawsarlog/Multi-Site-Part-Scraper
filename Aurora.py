import requests
import random
import csv
import pandas as pd
import re
import os
os.makedirs('output', exist_ok=True)
# all_proxies = ''
xlsx_file = "Parts for Xref 4.22.25.xlsx"  # Replace with your actual file name
Aurora_output = 'output/Aurora_output.csv'

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

    for item in data.get("content", []):
        if data_found == max_data:
            break

        sku = item.get("sku", "").strip().lower()
        title = " ".join(item.get("description", [])).strip().lower()
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
                    item.get('sku', 'N/A'),
                    item.get('description', ['N/A'])[0],
                    item.get('price', {}).get('amount', 'N/A'),
                    matching_res,
                    item.get('primaryAssetContentUrl', 'N/A')
                ]
            ])
            data_found += 1

    print("Collected Matches:", collect_data)
    return collect_data

def save_data(data):
    while len(data) < 18:
        data.append("N/A")
    with open(Aurora_output, mode='a', newline='',encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow([
        "Part Number Reformat", "DESCRIPTION", "Invoice Reference",
        "Aurora  Part # 1", "Aurora  Desc 1", "Aurora  Price 1", "Aurora  CS 1", "Aurora  Photo 1",
        "Aurora  Part # 2", "Aurora  Desc 2", "Aurora  Price 2", "Aurora  CS 2", "Aurora  Photo 2",
        "Aurora  Part # 3", "Aurora  Desc 3", "Aurora  Price 3", "Aurora  CS 3", "Aurora  Photo 3"
    ])
        writer.writerow(data)
        
def get_json_data(search_number):
    for _ in range(10):
        try:
            
            cookies = {
                'BLSR': 'eyJhbGciOiJSUzI1NiJ9.eyJpc3MiOiJicm9hZGxlYWYtYXV0aGVudGljYXRpb24iLCJzdWIiOiJCTFNSIiwiYXVkIjoiYnJvYWRsZWFmLWF1dGhlbnRpY2F0aW9uIiwicmVkaXJlY3RVcmwiOiJhSFIwY0hNNkx5OTNkM2N1WVhWeWIzSmhjR0Z5ZEhOMGIyZHZMbU52YlM5aGRYUm9MMjloZFhSb0wyRjFkR2h2Y21sNlpUOWpiR2xsYm5SZmFXUTlRVlZTVDFKQlgwTlZVMVJQVFVWU1gwTk1TVVZPVkNad2NtOXRjSFE5Ym05dVpTWnlaV1JwY21WamRGOTFjbWs5YUhSMGNITWxNMEVsTWtZbE1rWjNkM2N1WVhWeWIzSmhjR0Z5ZEhOMGIyZHZMbU52YlNVeVJuTnBiR1Z1ZEMxallXeHNZbUZqYXk1b2RHMXNKbkpsYzNCdmJuTmxYM1I1Y0dVOVkyOWtaU1p6WTI5d1pUMVZVMFZTSlRJd1ExVlRWRTlOUlZKZlZWTkZVaVp6ZEdGMFpUMWxlVXA2V1RJNWQxcFRTVFpKYkZaVVVsWkpaMUV4VmxSV1JUbE9VbFpLWmxaV1RrWlZhVWx6U1c1S2JGcHRWbmxqYlZaNVNXcHZhVXd6UW1oamJsRjBZekpXYUdOdFRtOVFNa3B6VVROV2VXTnRWblZaTTJzNVZsWk9SVWxwZDJsaWJUbDFXVEpWYVU5cFNYaE5SRnBvVFRKRk1VMVRNV3RPUjFaclRGUlJNRTFVWTNSUFIxRjZUMU14YUUxRVJteGFhbWhxV1ZkVk1Ga3lTV2xtVVNVelJDVXpSQT09IiwicmVxdWVzdFVybCI6ImFIUjBjSE02THk5M2QzY3VZWFZ5YjNKaGNHRnlkSE4wYjJkdkxtTnZiUzloZFhSb0wyOWhkWFJvTDJGMWRHaHZjbWw2WlE9PSJ9.k0pMNrozXygGtFSdfi3pKp9W53Ht2PIj5t9L75oTdfYI7tuxTzDqJmRLX5OgXJmTVIGjZ_O7OPLGOBknNSLbAY-JWCarnpzKAafMTftsYzCnL6t1I7YuRPvhkBRlLItzLntapsY8Kow_vZfB1IZ21Y0FZRPV7F4wlbscu08EGYaGvQJRYkI4GK0pLrpY1qQnTlw2nUYKUoZzyJDqxK1mD6IiypJSYH4bKS6FxZ0R-OsBb_pWKRPWYeXqmCm4deaC_UA682s-3414OKe_sbS-YAiVSl-eneKaYV-EWPISN3lZzZBgq7WSj5sxIFbOByWmtZrhsVQD4dxSD5dgiJh15g',
            }

            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:136.0) Gecko/20100101 Firefox/136.0',
                'Accept': 'application/json',
                'Accept-Language': 'en',
                'X-Context-Request': '{"applicationId": "AURORA", "customerContextId": "5DF1363059675161A85F576D", "sandboxId": null, "tenantId": "5DF1363059675161A85F576D"}',
                'X-Resolution-Context': '{"resolutionContext": "DEFAULT"}',
                'X-National-Site-Id': 'AURORA',
                'X-Price-Context': '{"locale": "en", "currency": "USD", "userTargets": [{"targetType": "CUSTOMER"}], "attributes": {"webRequest": {"fullUrl": "https://www.aurorapartstogo.com/search?q=3456", "pathname": "/search", "userAgentType": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:136.0) Gecko/20100101 Firefox/136.0", "secure": true, "searchKeywords": "?q=3456"}}}',
                'X-Price-Info-Context': '{"priceLists": [], "skipDetails": true}',
                'Connection': 'keep-alive',
                'Sec-Fetch-Dest': 'empty',
                'Sec-Fetch-Mode': 'cors',
                'Sec-Fetch-Site': 'same-origin',
            }
            
            params = {
                'page': '0',
                'query': search_number,
                'size': '1000',
                'sort': 'relevance',
                'type': 'PRODUCT',
            }

            response = requests.get(
                            'https://www.aurorapartstogo.com/api/catalog-browse/search/catalog',
                            params=params,
                            cookies=cookies,
                            headers=headers,
                            # proxies=proxies
                        )

            if response.status_code == 200:
                return response.json()
        except:
            pass
    return False

def collect_details(search_number, search_disc, d_col=None, e_col=None):

    # Skip invalid part numbers like "12:00"
    if re.match(r'^\d{1,2}:\d{2}$', search_number.strip()):
        print(f"Skipping invalid part number format: {search_number}")
        return False

    # Fetch data
    data = get_json_data(search_number)

    if not data:
        with open("error.txt", "a") as f:
            f.write(f"{search_number} - {search_disc}\n")
        return False

    total_item = data.get('totalElements', 0)
    print("Total found >>", total_item)

    # Get matches
    matches = find_best_matches(data, search_number, search_disc, d_col, e_col)
    matches = matches[:3]
    print("Page 1 - Best Matches:", len(matches))
    final_data = [search_number, search_disc, 'N/A']

    for m in matches:
        for n in m:
            final_data.append(n)  # SKU

    save_data(final_data)

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
    collect_details(search_number,search_disc,d_col,e_col)

print("Script Completed...")
