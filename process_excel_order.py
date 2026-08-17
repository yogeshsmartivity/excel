import os
import re
import sys
import argparse
import warnings
warnings.filterwarnings("ignore", category=FutureWarning)
import pandas as pd
import pypdf
import win32com.client

GITHUB_RAW_BASE = "https://raw.githubusercontent.com/yogeshsmartivity/excel/main/"
CURRENT_VERSION = "1.0.0"

def check_for_updates(workbook_path=None):
    """
    Checks GitHub Raw URL for online updates and silently downloads:
    1. process_excel_order.py (Backend Code)
    2. master_price_list.xlsx (Master Price List & Discounts)
    3. version.txt (Version Tracker)
    """
    try:
        import urllib.request
        wb_dir = os.path.dirname(os.path.abspath(workbook_path)) if workbook_path else os.path.dirname(os.path.abspath(__file__))
        version_file = os.path.join(wb_dir, "version.txt")
        
        current_ver = CURRENT_VERSION
        if os.path.exists(version_file):
            try:
                with open(version_file, "r", encoding="utf-8") as vf:
                    for line in vf:
                        if line.startswith("VERSION="):
                            current_ver = line.strip().split("=")[1]
            except Exception:
                pass
                
        print(f"System Version: v{current_ver} (GitHub Full Auto-Sync Enabled)")
        
        # Check GitHub raw version.txt
        try:
            online_ver_url = GITHUB_RAW_BASE + "version.txt"
            req = urllib.request.Request(online_ver_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=3) as resp:
                online_text = resp.read().decode('utf-8')
                online_ver = current_ver
                for l in online_text.split('\n'):
                    if l.startswith("VERSION="):
                        online_ver = l.strip().split("=")[1]
                        
                if online_ver != current_ver:
                    print(f"New update found on GitHub (v{current_ver} -> v{online_ver})! Auto-updating...")
                    
                    # Download updated process_excel_order.py
                    py_url = GITHUB_RAW_BASE + "process_excel_order.py"
                    py_path = os.path.join(wb_dir, "process_excel_order.py")
                    urllib.request.urlretrieve(py_url, py_path)
                    
                    # Download updated master_price_list.xlsx
                    price_url = GITHUB_RAW_BASE + "master_price_list.xlsx"
                    price_path = os.path.join(wb_dir, "master_price_list.xlsx")
                    urllib.request.urlretrieve(price_url, price_path)
                    
                    # Download updated master_discount_list.xlsx
                    disc_url = GITHUB_RAW_BASE + "master_discount_list.xlsx"
                    disc_path = os.path.join(wb_dir, "master_discount_list.xlsx")
                    try:
                        urllib.request.urlretrieve(disc_url, disc_path)
                    except Exception:
                        pass
                    
                    # Write updated version.txt
                    with open(version_file, "w", encoding="utf-8") as vf:
                        vf.write(f"VERSION={online_ver}\n")
                        
                    print(f"✅ System successfully auto-updated to v{online_ver} from GitHub!")
        except Exception:
            pass
            
    except Exception as e:
        print(f"Auto-update check skipped: {e}")

def sync_master_price_list(wb, wb_dir):
    """
    If master_price_list.xlsx or master_discount_list.xlsx exists, syncs updated rows into Price list and discount sheets.
    """
    master_path = os.path.join(wb_dir, "master_price_list.xlsx")
    if os.path.exists(master_path):
        try:
            print("Syncing Master Price List into active workbook...")
            import openpyxl
            wb_master = openpyxl.load_workbook(master_path, data_only=True)
            if "Price list" in wb_master.sheetnames:
                sh_m = wb_master["Price list"]
                sh_price = wb.Sheets("Price list")
                for r in range(2, min(500, sh_m.max_row + 1)):
                    std_sku = sh_m.cell(r, 2).value
                    if std_sku is not None:
                        for c in range(1, 10):
                            val = sh_m.cell(r, c).value
                            if val is not None:
                                try:
                                    sh_price.Cells(r, c).Value = str(val) if not isinstance(val, (int, float)) else val
                                except Exception:
                                    pass
                print("Price list synced successfully from Master File!")
        except Exception as sync_err:
            print(f"Warning syncing master price list: {sync_err}")

    # Sync Master Discount List
    disc_master_path = os.path.join(wb_dir, "master_discount_list.xlsx")
    target_disc_path = disc_master_path if os.path.exists(disc_master_path) else master_path
    if os.path.exists(target_disc_path):
        try:
            import openpyxl
            wb_disc = openpyxl.load_workbook(target_disc_path, data_only=True)
            if "discount" in wb_disc.sheetnames:
                print("Syncing Master Discount List into active workbook...")
                sh_dm = wb_disc["discount"]
                sh_disc_wb = wb.Sheets("discount")
                for r in range(2, min(200, sh_dm.max_row + 1)):
                    pname = sh_dm.cell(r, 1).value
                    if pname is not None:
                        for c in range(1, min(7, sh_dm.max_column + 1)):
                            val = sh_dm.cell(r, c).value
                            if val is not None:
                                try:
                                    sh_disc_wb.Cells(r, c).Value = str(val) if not isinstance(val, (int, float)) else val
                                except Exception:
                                    pass
                print("Discount list synced successfully from Master File!")
        except Exception as d_err:
            print(f"Warning syncing discount list: {d_err}")

def parse_pdf_with_gemini(file_path, api_key):
    import google.generativeai as genai
    import json
    
    print("Initializing Gemini API for AI-powered OCR...")
    genai.configure(api_key=api_key)
    
    # Upload the PDF file to Gemini API
    print(f"Uploading '{os.path.basename(file_path)}' to Gemini...")
    uploaded_file = genai.upload_file(file_path)
    
    # Use gemini-1.5-flash which is fast, free-tier supported, and natively handles PDF files.
    model = genai.GenerativeModel("gemini-1.5-flash")
    
    is_firstcry = "firstcry" in os.path.basename(file_path).lower()
    
    if is_firstcry:
        prompt = """
        Analyze this purchase order PDF.
        Extract the following information:
        1. The name of the buyer/party (e.g. DIGITAL AGE RETAIL PRIVATE LIMITED, etc.). Return this as 'party_name'.
        2. The list of items ordered. For each item, extract:
           - 'sku': The Product ID (e.g. 15433823).
           - 'name': The name/description of the item.
           - 'qty': The main ordered quantity (integer).
           - 'scheme': The free or scheme quantity if mentioned (integer, default 0).
           - 'mrp': The printed unit rate or MRP (float).
           
        Return the result strictly as a JSON object with two keys:
        {
          "party_name": "...",
          "items": [
            {
              "sku": "...",
              "name": "...",
              "qty": 10,
              "scheme": 0,
              "mrp": 100.0
            },
            ...
          ]
        }
        Make sure to output ONLY the raw JSON string. Do not wrap it in markdown backticks or blockquotes.
        """
    else:
        prompt = """
        Analyze this purchase order PDF.
        Extract the following information:
        1. The name of the buyer/party (e.g. APEX ENTERPRISES, BLINK COMMERCE PRIVATE LIMITED, etc.). Return this as 'party_name'.
        2. The list of items ordered. For each item, extract:
           - 'sku': The EAN/UPC barcode (13-digit number if visible) or the SKU code.
           - 'name': The name/description of the item.
           - 'qty': The main ordered quantity (integer).
           - 'scheme': The free or scheme quantity if mentioned (integer, default 0).
           - 'mrp': The printed unit rate or MRP (float).
           
        Return the result strictly as a JSON object with two keys:
        {
          "party_name": "...",
          "items": [
            {
              "sku": "...",
              "name": "...",
              "qty": 10,
              "scheme": 0,
              "mrp": 100.0
            },
            ...
          ]
        }
        Make sure to output ONLY the raw JSON string. Do not wrap it in markdown backticks or blockquotes.
        """
    
    print("Extracting fields using Gemini AI...")
    response = model.generate_content([uploaded_file, prompt])
    
    # Clean up the file from Gemini
    try:
        uploaded_file.delete()
    except Exception as del_err:
        print(f"VBA API Clean-up Warning: {del_err}")
        
    text = response.text.strip()
    # Clean output backticks if any
    if text.startswith("```json"):
        text = text[7:]
    elif text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    text = text.strip()
    
    data = json.loads(text)
    party_name = data.get("party_name", "UNKNOWN PARTY").strip()
    extracted_items = data.get("items", [])
    
    # Make sure all elements in extracted_items have correct types
    formatted_items = []
    for item in extracted_items:
        sku_val = str(item.get("sku", "")).strip()
        if sku_val.endswith(".0"):
            sku_val = sku_val[:-2]
            
        formatted_items.append({
            'sku': sku_val,
            'name': str(item.get("name", "")).strip(),
            'qty': int(item.get("qty", 0)),
            'scheme': int(item.get("scheme", 0)),
            'mrp': float(item.get("mrp", 0.0)),
            'base_cost': float(item.get("base_cost", 0.0))
        })
        
    return party_name, formatted_items

def parse_pdf_order(file_path):
    print(f"Parsing PDF Order: {os.path.basename(file_path)}...")
    reader = pypdf.PdfReader(file_path)
    lines = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            lines.extend(text.split('\n'))
            
    party_name = "UNKNOWN PARTY"
    for idx, l in enumerate(lines):
        if "BILLTO" in l.upper().replace(" ", ""):
            if idx + 1 < len(lines):
                party_name = lines[idx + 1].strip()
                break
                
    extracted_items = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        match_idx = re.match(r'^(\d+)(?:\s+(.*))?$', line)
        if match_idx:
            try:
                sr_num = int(match_idx.group(1))
                if sr_num > 100 or sr_num < 1:
                    i += 1
                    continue
            except ValueError:
                pass
                
            first_name_part = match_idx.group(2) if match_idx.group(2) else ""
            if any(k in first_name_part.upper() for k in ["REMARKS", "PENDING INVOICE", "FOC SCHEME", "INVOICE NO"]):
                i += 1
                continue
            
            # Find the SMRT line
            smrt_idx = -1
            for k in range(i, min(i + 8, len(lines))):
                if lines[k].strip().startswith("SMRT"):
                    smrt_idx = k
                    break
                    
            if smrt_idx != -1:
                # Reconstruct product name
                name_parts = []
                if first_name_part.strip() and not any(k in first_name_part.upper() for k in ["FOC SCHEME", "PENDING INVOICE"]):
                    name_parts.append(first_name_part.strip())
                for k in range(i + 1, smrt_idx):
                    part = lines[k].strip()
                    if not part.startswith("SMRT") and not part.isdigit() and "BILL TO" not in part and "ORDER INFO" not in part and "FOC SCHEME" not in part and "PENDING INVOICE" not in part:
                        name_parts.append(part)
                name = " ".join(name_parts)
                
                # Get the SMRT line tokens
                smrt_line = lines[smrt_idx].strip()
                tokens = smrt_line.split()
                
                # If numbers are on the next line, merge them
                is_merged = False
                if len(tokens) < 6 and smrt_idx + 1 < len(lines):
                    next_line = lines[smrt_idx + 1].strip()
                    smrt_line = smrt_line + " " + next_line
                    tokens = smrt_line.split()
                    is_merged = True
                    
                if len(tokens) >= 4:
                    # Find the first token starting with or containing ₹
                    mrp_idx = -1
                    for t_idx, token in enumerate(tokens):
                        if '₹' in token or token.startswith('₹') or (token.replace(',', '').replace('.', '').replace('₹','').isdigit() and t_idx > 0):
                            if '₹' in token or token.startswith('₹') or (token.replace(',', '').replace('.', '').isdigit() and t_idx == 1):
                                mrp_idx = t_idx
                                break
                                
                    # Fallback to first numeric token as MRP if no ₹ symbol is present
                    if mrp_idx == -1:
                        for t_idx, token in enumerate(tokens):
                            if t_idx > 0 and token.replace(',', '').replace('.', '').isdigit():
                                mrp_idx = t_idx
                                break
                                
                    if mrp_idx != -1 and mrp_idx + 2 < len(tokens):
                        sku = " ".join(tokens[0:mrp_idx])
                        mrp_str = tokens[mrp_idx].replace('₹', '').replace(',', '').strip()
                        try:
                            mrp = float(mrp_str)
                        except ValueError:
                            mrp = 0.0
                        try:
                            qty = int(tokens[mrp_idx + 1])
                            scheme = int(tokens[mrp_idx + 2])
                        except ValueError:
                            qty = 0
                            scheme = 0
                    else:
                        sku = tokens[0]
                        mrp_str = tokens[1].replace('₹', '').replace(',', '').strip()
                        try:
                            mrp = float(mrp_str)
                        except ValueError:
                            mrp = 0.0
                        try:
                            qty = int(tokens[2])
                            scheme = int(tokens[3])
                        except ValueError:
                            qty = 0
                            scheme = 0
                            
                    extracted_items.append({
                        'name': name,
                        'sku': sku,
                        'mrp': mrp,
                        'qty': qty,
                        'scheme': scheme
                    })
                    i = smrt_idx + (2 if is_merged else 1)
                    continue
        i += 1
    return party_name, extracted_items

def parse_blinkit_po(file_path):
    print(f"Parsing Blinkit PDF Order: {os.path.basename(file_path)}...")
    reader = pypdf.PdfReader(file_path)
    lines = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            for l in text.split('\n'):
                if l.strip():
                    lines.append(l.strip())
                    
    item_indices = [idx for idx, line in enumerate(lines) if line == "890801"]
    extracted_items = []
    
    for item_idx in item_indices:
        if item_idx + 2 >= len(lines):
            continue
        upc = "890801" + lines[item_idx+1] + lines[item_idx+2]
        
        # Extract description
        desc_parts = []
        curr_idx = item_idx + 3
        while curr_idx < len(lines):
            line = lines[curr_idx]
            if re.match(r'^\d+(\.\d+)?$', line):
                break
            desc_parts.append(line)
            curr_idx += 1
        description = " ".join(desc_parts)
        
        # Extract numbers
        numbers = []
        while curr_idx < len(lines):
            line = lines[curr_idx]
            if re.match(r'^\d+(\.\d+)?$', line) or line == ".":
                if line != ".":
                    numbers.append(line)
                curr_idx += 1
            else:
                break
                
        # Reconstruct split floats
        merged_nums = []
        n_idx = 0
        while n_idx < len(numbers):
            num = numbers[n_idx]
            if n_idx + 1 < len(numbers) and '.' in num:
                next_num = numbers[n_idx + 1]
                if len(next_num) == 1 and next_num.isdigit():
                    merged_nums.append(num + next_num)
                    n_idx += 2
                    continue
            merged_nums.append(num)
            n_idx += 1
            
        # Find Qty and MRP
        margin_idx = -1
        for j, num in enumerate(merged_nums):
            try:
                val = float(num)
                if abs(val - 45.0) < 0.1:
                    margin_idx = j
                    break
            except ValueError:
                pass
                
        if margin_idx != -1:
            try:
                qty = int(float(merged_nums[margin_idx - 2]))
                mrp = float(merged_nums[margin_idx - 1])
                extracted_items.append({
                    'name': description,
                    'sku': upc, # Temp store UPC in sku
                    'mrp': mrp,
                    'qty': qty,
                    'scheme': 0
                })
            except Exception as ex:
                print(f"Warning extracting fields: {ex}")
                
    party_name = "BLINK COMMERCE PRIVATE LIMITED"
    return party_name, extracted_items

def parse_firstcry_po(file_path):
    print(f"Parsing Firstcry PDF Order: {os.path.basename(file_path)}...")
    reader = pypdf.PdfReader(file_path)
    lines = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            for l in text.split('\n'):
                if l.strip():
                    lines.append(l.strip())
                    
    extracted_items = []
    i = 0
    expected_sr = 1
    
    while i < len(lines):
        line = lines[i].strip()
        if line.isdigit() and int(line) == expected_sr:
            style_idx = -1
            for j in range(i + 1, min(i + 25, len(lines))):
                if re.match(r'^SMRT\d+', lines[j]):
                    style_idx = j
                    break
            
            if style_idx != -1:
                style_code = lines[style_idx].strip()
                hsn = lines[style_idx + 1].strip() if style_idx + 1 < len(lines) else ""
                qty_str = lines[style_idx + 2].strip() if style_idx + 2 < len(lines) else "0"
                mrp_str = lines[style_idx + 3].strip() if style_idx + 3 < len(lines) else "0.0"
                
                product_id = ""
                for j in range(i + 1, style_idx):
                    val = lines[j].strip()
                    if val.isdigit() and 7 <= len(val) <= 9:
                        product_id = val
                        break
                
                desc_parts = []
                prod_id_idx = -1
                for j in range(i + 1, style_idx):
                    val = lines[j].strip()
                    if val == product_id:
                        prod_id_idx = j
                        break
                
                if prod_id_idx != -1:
                    for j in range(prod_id_idx + 1, style_idx):
                        desc_parts.append(lines[j].strip())
                description = " ".join(desc_parts).replace('\t', ' ').strip()
                
                try:
                    qty = int(qty_str)
                    mrp_clean = mrp_str.replace('₹', '').replace(',', '').strip()
                    mrp = float(mrp_clean)
                    
                    base_cost_str = lines[style_idx + 6].strip() if style_idx + 6 < len(lines) else "0.0"
                    base_cost_clean = base_cost_str.replace('₹', '').replace(',', '').strip()
                    base_cost = float(base_cost_clean)
                    
                    extracted_items.append({
                        'name': description,
                        'sku': product_id,
                        'mrp': mrp,
                        'qty': qty,
                        'scheme': 0,
                        'base_cost': base_cost
                    })
                    expected_sr += 1
                    i = style_idx + 4
                    continue
                except Exception as ex:
                    print(f"Error parsing item {expected_sr} near line {i}: {ex}")
        i += 1
        
    party_name = "Digital Age Retail Pvt. Ltd."
    return party_name, extracted_items

def parse_excel_order(file_path):
    print(f"Parsing Excel Order: {os.path.basename(file_path)}...")
    df = pd.read_excel(file_path)
    
    col_map = {col: str(col).strip().upper() for col in df.columns}
    df = df.rename(columns=col_map)
    
    sku_col = None
    qty_col = None
    scheme_col = None
    name_col = None
    mrp_col = None
    
    for original, clean in col_map.items():
        if clean in ['SKU', 'SKU CODE', 'ITEM CODE', 'ITEMCODE', 'BN ITEM NO', 'BNITEMNO', 'PRODUCT CODE']:
            sku_col = clean
        elif clean in ['QTY', 'QUANTITY', 'QTY ORDERED', 'ORDER QTY', 'QUANTITY ORDERED']:
            qty_col = clean
        elif clean in ['SCHEME', 'SCHEME QTY', 'FREE', 'FREE QTY', 'SCHEME QUANTITY']:
            scheme_col = clean
        elif clean in ['NAME', 'ITEM NAME', 'PRODUCT NAME', 'DESCRIPTION', 'PRODUCT DESCRIPTION']:
            name_col = clean
        elif clean in ['MRP', 'RATE', 'PRICE', 'UNIT PRICE', 'BASIC RATE', 'ORDER MRP']:
            mrp_col = clean
            
    if not sku_col and len(df.columns) > 0:
        sku_col = df.columns[0]
    if not qty_col and len(df.columns) > 1:
        qty_col = df.columns[1]
    if not scheme_col and len(df.columns) > 2:
        scheme_col = df.columns[2]
        
    extracted_items = []
    party_name = "UNKNOWN PARTY"
    
    for col in df.columns:
        for idx in range(min(5, len(df))):
            val = str(df.loc[idx, col]).upper()
            if "BILL TO" in val or "PARTY" in val or "CUSTOMER" in val:
                if idx + 1 < len(df):
                    party_name = str(df.loc[idx + 1, col]).strip()
                    break
                    
    if party_name == "UNKNOWN PARTY":
        filename = os.path.basename(file_path)
        party_name = os.path.splitext(filename)[0].replace("Order", "").replace("order", "").strip()
        
    for idx, row in df.iterrows():
        sku = str(row[sku_col]).strip() if sku_col and pd.notna(row[sku_col]) else ""
        if not sku or sku.upper() == 'NAN' or sku == "":
            continue
            
        qty = int(row[qty_col]) if qty_col and pd.notna(row[qty_col]) else 0
        scheme = int(row[scheme_col]) if scheme_col and pd.notna(row[scheme_col]) else 0
        name = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else ""
        mrp = float(row[mrp_col]) if mrp_col and pd.notna(row[mrp_col]) else 0.0
        
        extracted_items.append({
            'sku': sku,
            'qty': qty,
            'scheme': scheme,
            'name': name,
            'mrp': mrp
        })
    return party_name, extracted_items

def get_active_workbook(workbook_path):
    excel = None
    try:
        excel = win32com.client.GetObject(Class="Excel.Application")
    except Exception:
        print("Excel is not running. Launching Excel Application...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        
    try:
        excel.DisplayAlerts = False
    except Exception:
        pass
        
    wb = None
    target_name = os.path.basename(workbook_path).lower()
    for w in excel.Workbooks:
        if w.Name.lower() == target_name or w.FullName.lower() == workbook_path.lower():
            wb = w
            break
            
    if not wb:
        if os.path.exists(workbook_path):
            print(f"Opening workbook: {workbook_path}")
            wb = excel.Workbooks.Open(os.path.abspath(workbook_path))
        else:
            raise FileNotFoundError(f"Workbook '{target_name}' not found at '{workbook_path}'.")
            
    return excel, wb

def load_ean_mappings(wb):
    mappings = {}
    try:
        sh_price = wb.Sheets("Price list")
        last_row = sh_price.Cells(sh_price.Rows.Count, "B").End(-4162).Row
        for r in range(2, last_row + 1):
            sku = str(sh_price.Cells(r, 2).Value).strip()
            ean = str(sh_price.Cells(r, 7).Value).strip() # Column G (7) is EAN
            if ean and ean != "None" and sku and sku != "None":
                # Strip decimals if Excel imported it as float
                if ean.endswith(".0"):
                    ean = ean[:-2]
                mappings[ean] = sku
    except Exception as e:
        print(f"Warning: Could not load EAN mappings from Price list: {e}")
    return mappings

def load_firstcry_mappings(wb, workbook_path=None):
    mappings = {}
    
    # 1. Try to load from active win32com object
    try:
        sh_price = wb.Sheets("Price list")
        last_row = sh_price.Cells(sh_price.Rows.Count, "B").End(-4162).Row
        for r in range(2, last_row + 1):
            sku = str(sh_price.Cells(r, 2).Value).strip()
            fc_val = sh_price.Cells(r, 8).Value
            if fc_val is not None:
                fc_id = str(fc_val).strip()
                if fc_id.endswith(".0"):
                    fc_id = fc_id[:-2]
                if fc_id and fc_id != "None" and fc_id != "#N/A" and fc_id != "#VALUE!" and sku and sku != "None":
                    mappings[fc_id] = sku
    except Exception as e:
        print(f"Warning: Could not load Firstcry mappings via win32com: {e}")
        
    # 2. Fall back to pandas reading from disk if no mappings loaded
    if not mappings and workbook_path and os.path.exists(workbook_path):
        print("VBA COM returned no mappings. Falling back to reading cached values from disk via pandas...")
        try:
            import pandas as pd
            df = pd.read_excel(workbook_path, sheet_name="Price list")
            sku_col = None
            fc_col = None
            for col in df.columns:
                col_name = str(col).strip().upper()
                if col_name == 'SKU':
                    sku_col = col
                elif col_name == 'FIRSTCRY':
                    fc_col = col
                    
            if sku_col and fc_col:
                for idx, row in df.iterrows():
                    sku = str(row[sku_col]).strip()
                    fc_val = row[fc_col]
                    if pd.notna(fc_val):
                        fc_id = str(fc_val).strip()
                        if fc_id.endswith(".0"):
                            fc_id = fc_id[:-2]
                        if fc_id and fc_id.lower() != 'nan' and fc_id != '#n/a' and fc_id != '#value!' and sku and sku.lower() != 'nan':
                            mappings[fc_id] = sku
        except Exception as pd_err:
            print(f"Warning: Could not load Firstcry mappings from disk file: {pd_err}")
            
    return mappings

def parse_swiggy_po(pdf_path):
    print(f"Parsing Swiggy PDF Order: {os.path.basename(pdf_path)}...")
    reader = pypdf.PdfReader(pdf_path)
    lines = []
    for page in reader.pages:
        txt = page.extract_text()
        if txt:
            lines.extend(txt.split('\n'))
            
    party_name = "CLOUDKART VENTURES PRIVATE LIMITED"
    for idx, l in enumerate(lines):
        if "BILLING ADDRESS" in l.upper():
            if idx + 1 < len(lines):
                party_name = lines[idx+1].strip()
                break
                
    extracted_items = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        match_start = re.match(r'^(\d+)\s+(\d+)\s+(.*)', line)
        if match_start:
            swiggy_code = match_start.group(2)
            name_start = match_start.group(3)
            
            summary_idx = -1
            for k in range(i, min(i + 15, len(lines))):
                k_line = lines[k].strip()
                tokens = k_line.split()
                if len(tokens) >= 10 and tokens[0].isdigit() and len(tokens[0]) >= 6:
                    summary_idx = k
                    break
                    
            if summary_idx != -1:
                desc_parts = [name_start]
                for k in range(i + 1, summary_idx):
                    part = lines[k].strip()
                    if not part.startswith('Colour:') and not part.startswith('Size:') and not part.startswith('Brand:'):
                        desc_parts.append(part)
                raw_name = ' '.join(desc_parts)
                cleaned_name = re.sub(r'\s*\d+\.\d+\s+pack.*', '', raw_name, flags=re.IGNORECASE).strip()
                
                sum_tokens = lines[summary_idx].strip().split()
                hsn = sum_tokens[0]
                qty = int(sum_tokens[1])
                mrp = float(sum_tokens[2])
                base_cost = float(sum_tokens[3])
                
                extracted_items.append({
                    'name': cleaned_name,
                    'sku': swiggy_code,
                    'mrp': mrp,
                    'qty': qty,
                    'scheme': 0,
                    'base_cost': base_cost,
                    'hsn': hsn
                })
                i = summary_idx + 1
                continue
        i += 1
        
    return party_name, extracted_items

def detect_vendor(order_path):
    filename = os.path.basename(order_path).lower()
    if "firstcry" in filename:
        return "Firstcry Order"
    if "blinkit" in filename:
        return "Blinkit Order"
    if "swiggy" in filename or "cloudkart" in filename:
        return "Swiggy Order"
    if "amit" in filename:
        return "Amit Order"
        
    # Read PDF text to detect by content keywords
    ext = os.path.splitext(order_path)[1].lower()
    if ext == '.pdf':
        try:
            reader = pypdf.PdfReader(order_path)
            if len(reader.pages) > 0:
                text = reader.pages[0].extract_text()
                if text:
                    text_upper = text.upper()
                    if "DIGITAL AGE RETAIL" in text_upper or "FIRSTCRY" in text_upper:
                        return "Firstcry Order"
                    if "BLINK COMMERCE" in text_upper or "BLINKIT" in text_upper:
                        return "Blinkit Order"
                    if "CLOUDKART" in text_upper or "SWIGGY" in text_upper:
                        return "Swiggy Order"
                    if "APEX ENTERPRISES" in text_upper or "SMARTIVITY LABS" in text_upper:
                        return "Amit Order"
        except Exception as e:
            print(f"Warning during vendor auto-detection: {e}")
            
    # Default fallback
    return "Amit Order"

def run_import(order_path, workbook_path):
    print(f"Running Import to workbook: {workbook_path}...")
    
    excel, wb = get_active_workbook(workbook_path)
    
    # Auto-detect vendor sheet
    target_sheet_name = detect_vendor(order_path)
    print(f"Auto-detected target sheet for order: '{target_sheet_name}'")
    active_sheet_name = target_sheet_name
    sh_order = wb.Sheets(active_sheet_name)
    try:
        sh_order.Activate()
    except Exception:
        pass
    
    # 1. Parse Order File
    ext = os.path.splitext(order_path)[1].lower()
    is_blinkit = (active_sheet_name == "Blinkit Order" or "blinkit" in os.path.basename(order_path).lower())
    is_firstcry = (active_sheet_name == "Firstcry Order" or "firstcry" in os.path.basename(order_path).lower())
    is_swiggy = (active_sheet_name == "Swiggy Order" or "swiggy" in os.path.basename(order_path).lower() or "cloudkart" in os.path.basename(order_path).lower() or "cvpl" in os.path.basename(order_path).lower())
    
    if ext == '.pdf':
        # Inspect PDF text content to auto-detect vendor format
        pdf_content_text = ""
        try:
            reader_check = pypdf.PdfReader(order_path)
            for p_check in reader_check.pages[:2]:
                t_check = p_check.extract_text()
                if t_check:
                    pdf_content_text += t_check.upper() + " "
        except Exception:
            pass

        if "CLOUDKART" in pdf_content_text or "CHCPO" in pdf_content_text or "SWIGGY" in pdf_content_text:
            is_swiggy = True
        elif "BLINK COMMERCE" in pdf_content_text or "BLINKIT" in pdf_content_text:
            is_blinkit = True
        elif "FIRSTCRY" in pdf_content_text or "DIGITAL AGE RETAIL" in pdf_content_text:
            is_firstcry = True

        parsed_via_local = False
        items = []
        party_name = "UNKNOWN PARTY"
        
        # Try local parsing first
        try:
            if is_blinkit:
                party_name, items = parse_blinkit_po(order_path)
            elif is_firstcry:
                party_name, items = parse_firstcry_po(order_path)
            elif is_swiggy:
                party_name, items = parse_swiggy_po(order_path)
            else:
                party_name, items = parse_pdf_order(order_path)
            
            if items:
                parsed_via_local = True
                print("Successfully parsed PDF using local parser!")
        except Exception as local_err:
            print(f"Local parser failed: {local_err}")
            
        # Fallback to Gemini if local parsing failed
        if not parsed_via_local:
            print("Local parser failed. Checking if Gemini API can parse it...")
            wb_dir = os.path.dirname(os.path.abspath(workbook_path))
            key_file_path = os.path.join(wb_dir, "gemini_key.txt")
            api_key = None
            if os.path.exists(key_file_path):
                try:
                    with open(key_file_path, "r", encoding="utf-8") as kf:
                        api_key = kf.read().strip()
                except Exception as k_err:
                    print(f"Warning: Could not read gemini_key.txt: {k_err}")
                    
            if api_key and len(api_key) > 10 and not api_key.startswith("YOUR_GEMINI_API_KEY"):
                try:
                    party_name, items = parse_pdf_with_gemini(order_path, api_key)
                    if items:
                        print("Successfully parsed PDF using Gemini AI!")
                        # Write warning file to trigger Excel popup
                        warning_file_path = os.path.join(wb_dir, "format_changed_warning.txt")
                        try:
                            with open(warning_file_path, "w", encoding="utf-8") as wf:
                                wf.write("Local parser failed. Gemini AI was used to parse the file.")
                            print(f"Warning file created at: {warning_file_path}")
                        except Exception as wf_err:
                            print(f"Warning: Could not write warning file: {wf_err}")
                    else:
                        raise ValueError("Gemini returned zero items.")
                except Exception as gem_err:
                    print(f"Gemini API execution failed: {gem_err}")
                    raise RuntimeError("Failed to parse the PDF using both local parsers and Gemini AI.")
            else:
                raise RuntimeError("Local parser failed and no valid Gemini API key was found.")
    elif ext in ['.xls', '.xlsx']:
        party_name, items = parse_excel_order(order_path)
    else:
        print(f"Error: Unsupported file format '{ext}'")
        return
        
    print(f"Extracted Party Name: '{party_name}'")
    print(f"Extracted {len(items)} items.")
    
    if not items:
        print("Error: No items found in order file.")
        return
        
    # Load Price list references for Fuzzy SKU Match and Auto Price Correction
    import difflib
    wb_dir = os.path.dirname(os.path.abspath(workbook_path))
    sync_master_price_list(wb, wb_dir)
    
    print("Loading Price list references for smart matching and correction...")
    sh_price = wb.Sheets("Price list")
    last_price_row = max(sh_price.Cells(sh_price.Rows.Count, col).End(-4162).Row for col in ["B", "F", "G", "H", "I"])
    
    id_to_std_sku = {}
    base_to_std_sku = {}
    sku_prices = {}
    
    for r in range(2, last_price_row + 1):
        std_sku = str(sh_price.Cells(r, 2).Value).strip()
        if not std_sku or std_sku.lower() == "none":
            continue
            
        clean_sku = str(sh_price.Cells(r, 6).Value).strip().upper().replace(" ", "")
        ean = str(sh_price.Cells(r, 7).Value).strip().upper().replace(" ", "")
        fc_id = str(sh_price.Cells(r, 8).Value).strip().upper().replace(" ", "")
        swiggy_val = sh_price.Cells(r, 9).Value
        swiggy_id = str(swiggy_val).strip().upper().replace(" ", "") if swiggy_val is not None else ""
        
        if ean.endswith(".0"):
            ean = ean[:-2]
        if fc_id.endswith(".0"):
            fc_id = fc_id[:-2]
        if swiggy_id.endswith(".0"):
            swiggy_id = swiggy_id[:-2]
            
        price_val = sh_price.Cells(r, 3).Value
        try:
            sku_prices[std_sku] = float(price_val) if price_val is not None else 0.0
        except ValueError:
            sku_prices[std_sku] = 0.0
            
        id_to_std_sku[std_sku.upper().replace(" ", "")] = std_sku
        if clean_sku and clean_sku != "NONE":
            id_to_std_sku[clean_sku] = std_sku
            if clean_sku.startswith("SMRT") and len(clean_sku) >= 8:
                base_to_std_sku[clean_sku[:8]] = std_sku
        if ean and ean != "NONE":
            id_to_std_sku[ean] = std_sku
        if fc_id and fc_id != "NONE" and fc_id != "#N/A":
            id_to_std_sku[fc_id] = std_sku
        if swiggy_id and swiggy_id != "NONE" and swiggy_id != "#N/A":
            id_to_std_sku[swiggy_id] = std_sku
            
    # Resolve, Base Match and Correct Prices for all items
    for item in items:
        raw_sku = str(item['sku']).strip()
        if raw_sku.endswith(".0"):
            raw_sku = raw_sku[:-2]
            
        clean_key = raw_sku.upper().replace(" ", "")
        resolved_sku = None
        auto_matched = False
        
        if clean_key in id_to_std_sku:
            resolved_sku = id_to_std_sku[clean_key]
        elif clean_key.startswith("SMRT") and len(clean_key) >= 8:
            base_key = clean_key[:8]
            if base_key in base_to_std_sku:
                resolved_sku = base_to_std_sku[base_key]
                auto_matched = True
                print(f"Base SKU matched '{raw_sku}' to '{resolved_sku}'")
            else:
                resolved_sku = raw_sku
        else:
            resolved_sku = raw_sku
                
        item['sku'] = resolved_sku
        item['auto_matched_sku'] = auto_matched
        item['original_sku_po'] = raw_sku
        
        original_price = item['mrp']
        price_corrected = False
        
        if resolved_sku in sku_prices:
            list_price = sku_prices[resolved_sku]
            if list_price > 0.0 and abs(original_price - list_price) > 0.1:
                item['mrp'] = list_price
                price_corrected = True
                print(f"Price Auto-Corrected for {resolved_sku} (PO: {original_price} -> List: {list_price})")
                
        item['original_mrp_po'] = original_price
        item['price_corrected'] = price_corrected
            
    # 3. Clear existing table
    print(f"Clearing old data on {active_sheet_name} sheet...")
    sh_order.Range("C5").Value = ""
    sh_order.Range("C6").Value = ""
    
    last_row = sh_order.Cells(sh_order.Rows.Count, "A").End(-4162).Row # -4162 = xlUp
    if last_row >= 11:
        sh_order.Range(f"A11:D{last_row}").ClearContents()
        sh_order.Range(f"E11:M{last_row}").ClearContents()
        
    # 4. Populate table and write formulas dynamically
    print("Writing new items and formulas...")
    sh_order.Range("C5").Value = party_name
    sh_order.Range("C6").Value = os.path.basename(order_path)
    
    # Load Party Discounts for smart matching
    sh_disc = wb.Sheets("discount")
    last_disc_row = sh_disc.Cells(sh_disc.Rows.Count, "A").End(-4162).Row
    party_discounts = {}
    for dr in range(2, last_disc_row + 1):
        pname = str(sh_disc.Cells(dr, 1).Value or "").strip().upper()
        pclean = str(sh_disc.Cells(dr, 2).Value or "").strip().upper()
        d5 = float(sh_disc.Cells(dr, 3).Value or 0.0)
        d12 = float(sh_disc.Cells(dr, 4).Value or 0.0)
        d18 = float(sh_disc.Cells(dr, 5).Value or 0.0)
        d28 = float(sh_disc.Cells(dr, 6).Value or 0.0)
        disc_dict = {0.05: d5, 0.12: d12, 0.18: d18, 0.28: d28}
        if pname and pname != "NONE":
            party_discounts[pname] = disc_dict
        if pclean and pclean != "NONE":
            party_discounts[pclean] = disc_dict

    # Find matching party discount dict
    matched_party_disc = None
    clean_party_upper = party_name.upper().strip()
    for pk, pd_dict in party_discounts.items():
        if pk in clean_party_upper or clean_party_upper in pk:
            matched_party_disc = pd_dict
            break

    for idx, item in enumerate(items):
        r = 11 + idx
        sh_order.Cells(r, 1).Value = item['sku']
        sh_order.Cells(r, 2).Value = item['name']
        sh_order.Cells(r, 3).Value = item['qty']
        sh_order.Cells(r, 4).Value = item['scheme']
        sh_order.Cells(r, 5).Value = item['mrp']
        
        # Write Excel formulas dynamically
        match_term = f"LEFT(UPPER(SUBSTITUTE(A{r}, \" \", \"\")), 8) & \"*\""
        sh_order.Cells(r, 6).Value = f'=IF(ISBLANK(A{r}), "", IFERROR(INDEX(\'Price list\'!B:B, MATCH({match_term}, \'Price list\'!F:F, 0)), ""))'
        sh_order.Cells(r, 7).Value = f'=IF(ISBLANK(A{r}), "", IFERROR(INDEX(\'Price list\'!C:C, MATCH({match_term}, \'Price list\'!F:F, 0)), 0))'
        sh_order.Cells(r, 8).Value = f'=IF(ISBLANK(A{r}), "", IFERROR(INDEX(\'Price list\'!E:E, MATCH({match_term}, \'Price list\'!F:F, 0)), 0))'
        sh_order.Cells(r, 9).Value = f'=IF(ISBLANK(A{r}), "", IFERROR(INDEX(\'Price list\'!D:D, MATCH({match_term}, \'Price list\'!F:F, 0)), ""))'
        
        if active_sheet_name == "Firstcry Order":
            base_cost = item.get('base_cost', 0.0)
            sh_order.Cells(r, 10).Value = f'=IF(ISBLANK(A{r}), 0, IF(G{r}=0, 0, ROUND(1 - ({base_cost} / G{r}), 5)))'
        else:
            default_disc_val = 0.533898
            if matched_party_disc:
                default_disc_val = matched_party_disc.get(0.18, 0.533898)
            sh_order.Cells(r, 10).Value = f'=IF(ISBLANK(A{r}), "", IFERROR(IF(ROUND(H{r},2)=0.18, INDEX(discount!E:E, MATCH("*" & UPPER(TRIM($C$5)) & "*", discount!B:B, 0)), IF(ROUND(H{r},2)=0.12, INDEX(discount!D:D, MATCH("*" & UPPER(TRIM($C$5)) & "*", discount!B:B, 0)), IF(ROUND(H{r},2)=0.28, INDEX(discount!F:F, MATCH("*" & UPPER(TRIM($C$5)) & "*", discount!B:B, 0)), INDEX(discount!C:C, MATCH("*" & UPPER(TRIM($C$5)) & "*", discount!B:B, 0))))), {default_disc_val}))'
            
        sh_order.Cells(r, 11).Value = f'=IF(ISBLANK(A{r}), "", ROUND(G{r}*C{r}*J{r},2))'
        try:
            sh_order.Cells(r, 12).Value = f'=IF(ISBLANK(A{r}), "", (G{r}*C{r})-K{r})'
        except Exception:
            pass
        
        # Build status note based on fuzzy match & price correction
        status_note = ""
        if item.get('auto_matched_sku') and item.get('price_corrected'):
            status_note = f"✅ Auto-matched SKU & Price (PO: '{item['original_sku_po']}', \u20b9{item['original_mrp_po']:.2f})"
        elif item.get('auto_matched_sku'):
            status_note = f"✅ Auto-matched SKU (PO: '{item['original_sku_po']}')"
        elif item.get('price_corrected'):
            status_note = f"✅ Price Auto-Corrected (PO: \u20b9{item['original_mrp_po']:.2f})"
            
        if status_note:
            sh_order.Cells(r, 13).Value = status_note
        else:
            sh_order.Cells(r, 13).Value = f'=IF(ISBLANK(A{r}), "", IF(F{r}="", "⚠️ SKU not found in Price List", IF(ROUND(E{r},2)<>ROUND(G{r},2), "⚠️ Price Mismatch! (PO: " & TEXT(E{r},"#,##0.00") & ", List: " & TEXT(G{r},"#,##0.00") & ")", "✅ OK")))'
        
    try:
        excel.Calculate()
        print("Excel formulas recalculated successfully.")
    except Exception as calc_err:
        print(f"Warning: Could not trigger Excel calculation: {calc_err}")
        
    print("Import and verification completed successfully!")

def run_fill(workbook_path):
    print(f"Reading verified order rows from workbook: {workbook_path}...")
    excel, wb = get_active_workbook(workbook_path)
    
    active_sheet_name = None
    try:
        curr_name = excel.ActiveSheet.Name
        if curr_name in ["Amit Order", "Blinkit Order", "Firstcry Order", "Swiggy Order"] and wb.Sheets(curr_name).Range("C5").Value:
            active_sheet_name = curr_name
    except Exception:
        pass
        
    if not active_sheet_name:
        for sname in ["Amit Order", "Blinkit Order", "Firstcry Order", "Swiggy Order"]:
            if wb.Sheets(sname).Range("C5").Value:
                active_sheet_name = sname
                break
                
    if not active_sheet_name:
        active_sheet_name = "Amit Order"
        
    print(f"Reading from order sheet: '{active_sheet_name}'")
    sh_order = wb.Sheets(active_sheet_name)
    sh_temp = wb.Sheets("template")
    
    try:
        sh_order.Calculate()
        excel.Calculate()
        print("Excel formulas recalculated successfully.")
    except Exception as calc_err:
        print(f"Warning during recalculation: {calc_err}")
        
    party_name = sh_order.Range("C5").Value
    order_file = sh_order.Range("C6").Value
    if not party_name:
        print("Error: No order loaded in the sheet.")
        return
        
    # 1. Read Order Rows
    last_row = sh_order.Cells(sh_order.Rows.Count, "A").End(-4162).Row
    if last_row < 11:
        print("Error: No items found in the Order table.")
        return
        
    rows_to_fill = []
    print(f"Reading rows from row 11 to {last_row}...")
    
    for r in range(11, last_row + 1):
        sku_val = sh_order.Cells(r, 6).Value # Column F is Price List SKU
        sku = str(sku_val).strip() if sku_val is not None else ""
        if not sku or sku.upper() == 'NONE' or sku == "":
            sku_val = sh_order.Cells(r, 1).Value # Fallback to Column A (SMRT SKU)
            sku = str(sku_val).strip() if sku_val is not None else ""
            
        if not sku or sku.upper() == 'NONE' or sku == "":
            continue
            
        qty = int(sh_order.Cells(r, 3).Value) if sh_order.Cells(r, 3).Value else 0
        scheme = int(sh_order.Cells(r, 4).Value) if sh_order.Cells(r, 4).Value else 0
        
        # Read looked-up values
        rate = float(sh_order.Cells(r, 7).Value) if sh_order.Cells(r, 7).Value else 0.0
        tax = float(sh_order.Cells(r, 8).Value) if sh_order.Cells(r, 8).Value else 0.0
        discount_pct = float(sh_order.Cells(r, 10).Value) if sh_order.Cells(r, 10).Value else 0.0
        discount_amt = float(sh_order.Cells(r, 11).Value) if sh_order.Cells(r, 11).Value else 0.0
        
        rows_to_fill.append({
            'sku': sku,
            'qty': qty,
            'scheme': scheme,
            'rate': rate,
            'tax': tax,
            'discount_pct': discount_pct,
            'discount_amt': discount_amt
        })
        
    if active_sheet_name == "Firstcry Order":
        aggregated = {}
        for rinfo in rows_to_fill:
            key = (str(rinfo['sku']).strip(), round(rinfo['discount_pct'], 5))
            if key not in aggregated:
                aggregated[key] = {
                    'sku': rinfo['sku'],
                    'qty': 0,
                    'scheme': 0,
                    'rate': rinfo['rate'],
                    'tax': rinfo['tax'],
                    'discount_pct': rinfo['discount_pct'],
                    'discount_amt': 0.0
                }
            aggregated[key]['qty'] += rinfo['qty']
            aggregated[key]['scheme'] += rinfo['scheme']
            
        for agg_info in aggregated.values():
            agg_info['discount_amt'] = round(agg_info['rate'] * agg_info['qty'] * agg_info['discount_pct'], 2)
            
        rows_to_fill = list(aggregated.values())
        print(f"Aggregated Firstcry rows: {len(rows_to_fill)} unique rows.")
        
    if not rows_to_fill:
        print("Error: No valid rows to process.")
        return
        
    # 2. Clear Template Sheet
    print("Clearing template sheet...")
    last_temp_row = sh_temp.Cells(sh_temp.Rows.Count, "A").End(-4162).Row
    if last_temp_row >= 2:
        sh_temp.Range(f"A2:K{last_temp_row}").ClearContents()
        
    # 3. Build output rows (interleaving main and scheme rows)
    template_rows = []
    
    for item in rows_to_fill:
        # Add main item row
        if item['qty'] > 0:
            disc_pct = round(item['discount_pct'] * 100.0, 4) if item['discount_pct'] > 0 else 0.0
            
            template_rows.append([
                item['sku'],
                "", # VariantCodeNo
                item['qty'],
                item['rate'],
                disc_pct,
                "", # DiscountAmount (blank as requested)
                item['tax'] * 100.0 if item['tax'] > 0 else 0.0,
                "", # CustomerItemCode
                "", # CustomerItemName
                "", # ItemRemark
                ""  # InputField
            ])
            
        # Add scheme item row (placed immediately below main item row)
        if item['scheme'] > 0:
            rate_val = item['rate']
            scheme_qty = item['scheme']
            
            template_rows.append([
                item['sku'],
                "", # VariantCodeNo
                scheme_qty,
                rate_val, # Price is the base rate (not 0.0)
                100.0, # 100% discount
                "", # DiscountAmount (blank as requested)
                item['tax'] * 100.0 if item['tax'] > 0 else 0.0,
                "a", # CustomerItemCode
                "",
                "",
                ""
            ])
            
    # 4. Write to Template Sheet
    print(f"Writing {len(template_rows)} rows to template sheet...")
    for idx, row_val in enumerate(template_rows):
        r_temp = 2 + idx
        for col_idx, val in enumerate(row_val, 1):
            sh_temp.Cells(r_temp, col_idx).Value = val
            
    # 5. Auto-Export Template to separate xlsx workbook silently
    try:
        base_name = os.path.splitext(os.path.basename(str(order_file)))[0] if order_file else "Order"
        default_filename = f"{base_name}_mapped.xlsx"
        export_path = os.path.join(os.path.dirname(os.path.abspath(workbook_path)), default_filename)
        
        if os.path.exists(export_path):
            try:
                os.remove(export_path)
            except Exception:
                pass
                
        print(f"Exporting template silently to: {export_path}...")
        sh_temp.Copy()
        new_wb = excel.ActiveWorkbook
        new_wb.SaveAs(os.path.abspath(export_path), 51) # 51 = xlOpenXMLWorkbook (.xlsx)
        new_wb.Close(False)
        print("Template exported successfully!")
    except Exception as exp_err:
        print(f"Warning during auto-export: {exp_err}")
        
    try:
        wb.Save()
    except Exception:
        pass
        
    print("Template populated successfully!")

if __name__ == "__main__":
    import traceback
    
    args_workbook = None
    try:
        parser = argparse.ArgumentParser(description="Process Excel Order Sheets")
        parser.add_argument("--import", action="store_true", dest="import_mode", help="Run in order import mode")
        parser.add_argument("--fill", action="store_true", dest="fill_mode", help="Run in template fill mode")
        parser.add_argument("--order", type=str, help="Path to selected order file")
        parser.add_argument("--workbook", type=str, required=True, help="Path to active workbook")
        
        args = parser.parse_args()
        args_workbook = args.workbook
        
        check_for_updates(args_workbook)
        
        if args.import_mode:
            if not args.order:
                print("Error: --order is required in import mode.")
                sys.exit(1)
            run_import(args.order, args.workbook)
        elif args.fill_mode:
            run_fill(args.workbook)
        else:
            print("Error: Specify --import or --fill")
            sys.exit(1)
            
    except Exception as e:
        wb_dir = os.path.dirname(os.path.abspath(args_workbook)) if args_workbook else os.path.dirname(os.path.abspath(__file__))
        log_path = os.path.join(wb_dir, "error_log.txt")
        try:
            with open(log_path, "w", encoding="utf-8") as f:
                f.write("=== Python Process Error Log ===\n")
                f.write(traceback.format_exc())
            print(f"Exception occurred. Error written to: {log_path}")
        except Exception as write_err:
            print(f"Failed to write log file: {write_err}")
        sys.exit(1)
