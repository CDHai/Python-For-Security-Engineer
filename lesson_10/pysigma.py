import os
import pandas as pd
import yaml
import schedule
import time
import subprocess
from sigma.rule import SigmaRule
from sigma.backends.splunk import SplunkBackend

SIGMA_REPO_URL = "https://github.com/SigmaHQ/sigma.git"
SIGMA_REPO_DIR = "C:/Users/admin/Desktop/Python-For-Security-Engineer/lesson_10/sigma"
SIGMA_RULES_DIR = os.path.join(SIGMA_REPO_DIR, "rules", "windows", "create_remote_thread")
EXCEL_FILE = "C:/Users/admin/Desktop/Python-For-Security-Engineer/lesson_10/output.xlsx"

def pull_sigma_rules():
    if os.path.exists(SIGMA_REPO_DIR):
        subprocess.run(["git", "-C", SIGMA_REPO_DIR, "pull"], check=True)
    else:
        subprocess.run(["git", "clone", SIGMA_REPO_URL, SIGMA_REPO_DIR], check=True)

def convert_sigma_to_splunk(sigma_file_path):
    try:
        backend = SplunkBackend()
        rule = SigmaRule.from_yaml(sigma_file_path)
        return backend.convert(rule)
    
    except Exception as e:
        print(e)

def extract_sigma_info(sigma_file_path):
    try:
        with open(sigma_file_path, 'r', encoding='utf-8') as sigma_file:
            content = yaml.safe_load(sigma_file)
            title = content.get("title", "")
            description = content.get("description", "")
            technique = content.get("detection", {}).get("condition", "")
            splunk_query = convert_sigma_to_splunk(sigma_file_path)
            return title, description, technique, splunk_query
    except (UnicodeDecodeError, yaml.YAMLError) as e:
        print(f"Error reading {sigma_file_path}: {e}")
        return "", "", "", ""

def convert_sigma_to_excel():
    data = []
    for root, _, files in os.walk(SIGMA_RULES_DIR):
        for file in files:
            if file.endswith(".yml") or file.endswith(".yaml"):
                sigma_file_path = os.path.join(root, file)
                title, description, technique, splunk_query = extract_sigma_info(sigma_file_path)
                data.append({
                    "File Name": file,
                    "Title": title,
                    "Description": description,
                    "Technique": technique,
                    "Query": splunk_query
                })
    df = pd.DataFrame(data)
    df.to_excel(EXCEL_FILE, index=False)

def job():
    pull_sigma_rules()
    convert_sigma_to_excel()

def main():
    schedule.every(1).minutes.do(job)
    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == "__main__":
    main()
