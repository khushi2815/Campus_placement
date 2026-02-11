import os
import time
import uuid
import zipfile
import json
import pandas as pd
import requests
from flask import Flask, request, send_file, jsonify
from bs4 import BeautifulSoup
from PIL import Image
from io import BytesIO
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
PHOTO_FOLDER = 'photos'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PHOTO_FOLDER, exist_ok=True)

# --- SCRAPER UTILITIES ---

def get_leetcode_rank(url):
    try:
        # LC uses GraphQL/React; basic scraping might need a header or session
        response = requests.get(url, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')
        # Simplified: In a real scenario, use LC's public GraphQL endpoint
        rank_tag = soup.find(string=lambda x: "Global Rank" in x)
        return rank_tag.parent.find_next_sibling().text if rank_tag else "N/A"
    except Exception as e: return f"Error: {str(e)}"

def get_codeforces_rating(url):
    try:
        response = requests.get(url, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')
        rating_span = soup.find('span', style=lambda x: x and 'font-weight:bold' in x)
        return rating_span.text if rating_span else "N/A"
    except Exception as e: return f"Error: {str(e)}"

def get_github_stats(url):
    try:
        response = requests.get(url, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')
        commits = soup.find('h2', class_='f4 text-normal mb-2').text.strip().split()[0]
        repos = soup.find('span', class_='Counter', data_view_component="true").text.strip()
        return commits, repos
    except: return "N/A", "N/A"

def get_linkedin_photo(url, roll_no):
    # Requires Selenium for dynamic content
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(options=chrome_options)
    try:
        driver.get(url)
        time.sleep(3) # Wait for JS
        img = driver.find_element("xpath", "//img[contains(@class, 'pv-top-card-profile-picture')]")
        img_url = img.get_attribute("src")
        
        img_data = requests.get(img_url).content
        img_obj = Image.open(BytesIO(img_data))
        path = os.path.join(PHOTO_FOLDER, f"{roll_no}.jpg")
        img_obj.convert('RGB').save(path, "JPEG")
        return path
    except: return "N/A"
    finally: driver.quit()

# --- CORE PROCESSING LOGIC ---

@app.route('/enrich', methods=['POST'])
def enrich_excel():
    if 'excel' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files['excel']
    job_id = str(uuid.uuid4())
    input_path = os.path.join(UPLOAD_FOLDER, f"{job_id}.xlsx")
    file.save(input_path)

    # 1. Load Data
    df = pd.read_excel(input_path)
    logs = []
    summary = {"total_rows": len(df), "platforms": {}}

    # 2. Process Rows
    for index, row in df.iterrows():
        roll_no = row.get('RollNo', f"user_{index}")
        
        # LeetCode
        if 'LeetCodeURL' in df.columns:
            df.at[index, 'LC_Global_Contest_Rank'] = get_leetcode_rank(row['LeetCodeURL'])
        
        # Codeforces
        if 'CodeforcesURL' in df.columns:
            df.at[index, 'CF_Rating'] = get_codeforces_rating(row['CodeforcesURL'])

        # GitHub
        if 'GitHubURL' in df.columns:
            c, r = get_github_stats(row['GitHubURL'])
            df.at[index, 'GH_Commits_12mo'], df.at[index, 'GH_Public_Repos'] = c, r
            
        # LinkedIn
        if 'LinkedInURL' in df.columns:
            df.at[index, 'Photos_Path'] = get_linkedin_photo(row['LinkedInURL'], roll_no)

        # Basic Log
        logs.append({"row": index, "status": "Success", "timestamp": str(datetime.now())})
        time.sleep(1) # Simple rate limit

    # 3. Save Enriched Excel
    output_xlsx = os.path.join(UPLOAD_FOLDER, f"enriched_{job_id}.xlsx")
    with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Enriched_Data')
        pd.DataFrame(logs).to_excel(writer, index=False, sheet_name='Enrich_Logs')

    # 4. Generate Summary & ZIP
    summary_path = os.path.join(UPLOAD_FOLDER, f"summary_{job_id}.json")
    with open(summary_path, 'w') as f:
        json.dump(summary, f)

    zip_path = os.path.join(UPLOAD_FOLDER, f"result_{job_id}.zip")
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        zipf.write(output_xlsx, arcname="enriched.xlsx")
        zipf.write(summary_path, arcname="summary.json")

    return send_file(zip_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
