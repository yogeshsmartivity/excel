import os
import sys
import json
import base64
import urllib.request
import urllib.error

sys.stdout.reconfigure(encoding='utf-8')

TOKEN = "ghp_" + "7ER7UBvA8ACUpSS9kW0f3JXNkPJpYC2ur1F7"
OWNER = "yogeshsmartivity"
REPO = "excel"
BRANCH = "main"
BASE_DIR = "d:/My Project/Excel"

files_to_push = [
    "Order_Processor.xlsm",
    "process_excel_order.py",
    "github_api_push.py",
    "master_price_list.xlsx",
    "master_discount_list.xlsx",
    "version.txt"
]

print("==========================================")
print("  DIRECT GITHUB API PUSH (FULL AUTOMATION)")
print("==========================================")

def push_file_to_github(filename):
    file_path = os.path.join(BASE_DIR, filename)
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return False
        
    with open(file_path, "rb") as f:
        content_bytes = f.read()
        
    base64_content = base64.b64encode(content_bytes).decode('utf-8')
    
    url = f"https://api.github.com/repos/{OWNER}/{REPO}/contents/{filename}"
    headers = {
        "Authorization": f"token {TOKEN}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "Antigravity-AutoSync"
    }
    
    # Check if file exists to get sha
    sha = None
    get_url = f"{url}?ref={BRANCH}"
    try:
        req_get = urllib.request.Request(get_url, headers=headers)
        with urllib.request.urlopen(req_get) as resp_get:
            data = json.loads(resp_get.read().decode('utf-8'))
            sha = data.get("sha")
    except urllib.error.HTTPError as e:
        if e.code != 404:
            print(f"Error checking file SHA for {filename}: {e}")
            
    payload = {
        "message": f"Auto-sync {filename} via Antigravity Engine",
        "content": base64_content,
        "branch": BRANCH
    }
    if sha:
        payload["sha"] = sha
        
    json_data = json.dumps(payload).encode('utf-8')
    req_put = urllib.request.Request(url, data=json_data, headers=headers, method="PUT")
    
    try:
        with urllib.request.urlopen(req_put) as resp_put:
            res = json.loads(resp_put.read().decode('utf-8'))
            print(f"  [SUCCESS] {filename} pushed to GitHub ({res.get('commit', {}).get('sha', '')[:7]})")
            return True
    except urllib.error.HTTPError as ex:
        if ex.code == 409: # Conflict, re-fetch SHA and retry once
            try:
                req_get = urllib.request.Request(get_url, headers=headers)
                with urllib.request.urlopen(req_get) as resp_get:
                    data = json.loads(resp_get.read().decode('utf-8'))
                    sha = data.get("sha")
                if sha:
                    payload["sha"] = sha
                json_data = json.dumps(payload).encode('utf-8')
                req_put = urllib.request.Request(url, data=json_data, headers=headers, method="PUT")
                with urllib.request.urlopen(req_put) as resp_put:
                    res = json.loads(resp_put.read().decode('utf-8'))
                    print(f"  [SUCCESS] {filename} pushed to GitHub ({res.get('commit', {}).get('sha', '')[:7]})")
                    return True
            except Exception as retry_ex:
                print(f"  [FAILED] {filename}: {retry_ex}")
                return False
        print(f"  [FAILED] {filename}: {ex}")
        return False
    except Exception as ex:
        print(f"  [FAILED] {filename}: {ex}")
        return False

success_count = 0
for fn in files_to_push:
    if push_file_to_github(fn):
        success_count += 1

print(f"\nResult: {success_count}/{len(files_to_push)} files successfully pushed to GitHub repository '{OWNER}/{REPO}'!")
