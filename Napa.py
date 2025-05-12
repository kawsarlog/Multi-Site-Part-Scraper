import requests
import csv
import pandas as pd
import re
import os
os.makedirs('output', exist_ok=True)

xlsx_file = "Parts for Xref 4.22.25.xlsx"  # Replace with your actual file name

output_file_name = "output/napa_output.csv"

def normalize(text):
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
        
    with open(output_file_name, mode='a', newline='',encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow([
        "Part Number Reformat", "DESCRIPTION", "Invoice Reference",
        "NAPA Part # 1", "NAPA Desc 1", "NAPA Price 1", "NAPA CS 1", "Napa Photo 1",
        "NAPA Part # 2", "NAPA Desc 2", "NAPA Price 2", "NAPA CS 2", "Napa Photo 2",
        "NAPA Part # 3", "NAPA Desc 3", "NAPA Price 3", "NAPA CS 3", "Napa Photo 3"
    ])
        writer.writerow(data)

def get_json_data(search_number,start):
    for _ in range(10):
        try:
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:136.0) Gecko/20100101 Firefox/136.0',
                'Accept': '*/*',
                'Accept-Language': 'en-US,en;q=0.5',
                'Referer': 'https://www.napaonline.com/',
                'content-type': 'application/x-www-form-urlencoded',
                'x-gpc-clientid': 'NOL',
                'x-gpc-experimentid': 'QUS_TREATMENT',
                'x-gpc-requestid': 'd84d4751ec834686a78ab53318d2c1c2',
                'x-gpc-service': 'SEARCH',
                'x-gpc-storeid': 'Store900001025',
                'Origin': 'https://www.napaonline.com',
                'Connection': 'keep-alive',
                'Sec-Fetch-Dest': 'empty',
                'Sec-Fetch-Mode': 'cors',
                'Sec-Fetch-Site': 'cross-site',
                'Priority': 'u=4',
            }

            params = {
                'site': 'us',
                'fl': 'pid,title,brand,sale_price,primary_image,thumb_image,url,description,unit_of_measure,product_line,line_abbreviation,part_number,interchange_parts,universal,field_sku,hq_abbrev,regulatory,priced_from,priced_to,retail_redirect,retail_url,part_type',
                '_br_uid_2': '_br_uid_2_unavailable',
                'request_type': 'search',
                'search_type': 'keyword',
                'domain_key': 'napaonline',
                'url': f'https://www.napaonline.com/en/search?text={search_number}&referer=v2&page=1',
                'facet.range': 'price',
                'facet.application_part_type.limit': '300',
                'q': search_number,
                'view_id': 'FRE',
                'efq': 'dc:"FRE"',
                'rows': '120',
                'start': start,
                'sort': 'relevance asc',
            }

            response = requests.post(
                'https://api.genpt.com/discovery/api/v1/core/',
                headers=headers,
                params=params
            )

            if response.status_code == 200:
                return response.json()
            print(f"Status code : {response.status_code}")
        except:
            pass
    return False

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

    for item in data.get("response", {}).get("docs", []):
        if data_found == max_data:
            break

        sku = item.get("field_sku", "").strip().lower()
        part_number = item.get("part_number", "").strip().lower()
        title = item.get("title", "").strip().lower()
        if not sku or not title: 
            continue

        sku_n, sku_c = normalize(sku), clean_and_lower(sku)
        part_number_n, part_number_c = normalize(part_number), clean_and_lower(part_number)
        title_c = clean_and_lower(title)

        interchange_parts = item.get("interchange_parts", [])
        interchange_parts_n = [normalize(p) for p in interchange_parts]
        interchange_parts_c = [clean_and_lower(p) for p in interchange_parts]

        matched = False
        matching_res = ""
        exact_conditions = [
                                    e_col_n == sku_n, 
                                    e_col_c == sku_c,
                                    d_col_n == sku_n, 
                                    d_col_c == sku_c,
                                    search_number_n == sku_n, 
                                    search_number_c == sku_c
                                ]
        partial_conditions = [
                                    e_col_n in sku_n, 
                                    e_col_c in sku_c,
                                    d_col_n in sku_n, 
                                    d_col_c in sku_c,
                                    search_number_n in sku_n, 
                                    search_number_c in sku_c
                                ]
        # ✅ Exact Match
        # exact_conditions = [
        #     e_col_n in [sku_n, part_number_n], e_col_c in [sku_c, part_number_c],
        #     search_number_n in [sku_n, part_number_n], search_number_c in [sku_c, part_number_c]
        # ]
        if any(exact_conditions) and (word_exists(title_c, search_disc_c) or word_exists(title_c, d_col_c)):
            matching_res = 'Exact Match'
            matched = True
            print(f"[Exact] Matched SKU/Part#: {sku} | Title: {title}")

        # ✅ x-ref match
        # elif any([
        #     word_ex c), word_exists(e_col_c, part_number_c),
        #     word_exists(search_number_c, sku_c), word_exists(search_number_c, part_number_c)
        # ]) and (word_exists(title_c, search_disc_c) or word_exists(title_c, d_col_c)):
        elif any(partial_conditions) and not search_number_c.isdigit():
            matching_res = 'x-ref match'
            matched = True
            print(f"[x-ref] Substring Match with SKU/Part#: {sku} | Title: {title}")

        # ✅ Interchange Match
        elif any(x in [sku_c, part_number_c, search_number_c, e_col_c] for x in interchange_parts_c) and \
            (word_exists(title_c, search_disc_c) or word_exists(title_c, d_col_c)):
            matching_res = 'x-ref match'
            matched = True
            print(f"[Interchange C] Matched Interchange Part (cleaned) in: {interchange_parts}")

        # elif any(x in [sku_n, part_number_n, search_number_n, e_col_n] for x in interchange_parts_n) and \
        #     (word_exists(title_c, search_disc_c) or word_exists(title_c, d_col_c)):
        #     matching_res = 'x-ref match'
        #     matched = True
        #     print(f"[Interchange N] Matched Interchange Part (normalized) in: {interchange_parts}")

        if matched:
            collect_data.extend([
            [
                item.get('field_sku', 'N/A'),
                item.get('title', 'N/A'),
                item.get('sale_price', 'N/A'),
                matching_res,
                item.get('thumb_image', 'N/A')
            ]
        ])
            data_found += 1

    print("Collected Matches:", collect_data)
    return collect_data

def collect_details(search_number, search_disc, d_col, e_col):
    start = 0
    data = get_json_data(search_number, start)

    if data is False:
        with open("error.txt", "a") as f:
            f.write(f"{search_number} - {search_disc}\n")
        return False

    total_items = data.get('response', {}).get('numFound', 0)
    total_pages = (total_items + 120 - 1) // 120
    print("Total Pages:", total_pages)

    best_matches = []

    for page in range(total_pages):
        if len(best_matches) >= 3:
            break

        if page != 0:
            start = page * 120
            data = get_json_data(search_number, start)
            if data is False:
                with open("error.txt", "a") as f:
                    f.write(f"{search_number} - {search_disc}\n")
                return False

        matches = find_best_matches(data, search_number, search_disc, d_col, e_col)
        best_matches.extend(matches)

        # Keep only the top 3 matches
        best_matches = best_matches[:3]
        print(f"Page {page + 1} - Best Matches: {len(best_matches)}")
        if len(best_matches) == 3:
            break


    # Build the final data row
    final_data = [search_number, search_disc, 'N/A']

    for m in best_matches:
        for n in m:
            final_data.append(n)  # SKU

    save_data(final_data)

df = pd.read_excel(xlsx_file, engine="openpyxl")
df_filtered = df[1:]


for index, row in df_filtered.iterrows():
    print(f"\n>>>>> : {index + 1}")
    search_number = row.tolist()[1]
    search_number = str(search_number).strip() if search_number is not None else None

    search_disc = row.tolist()[2]
    search_disc = str(search_disc).strip() if search_disc is not None else None

    d_col = row.tolist()[4]
    d_col = str(d_col).strip() if d_col is not None else ''

    e_col = row.tolist()[5]
    e_col = str(e_col).strip() if e_col is not None else ''

    print(search_number," >>> ", search_disc)
    collect_details(search_number,search_disc,d_col,e_col)

print("Script Completed...")
