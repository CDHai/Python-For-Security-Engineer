import subprocess
import os
import pandas as pd
import schedule
import time

SIGMA_REPO_URL = "https://github.com/SigmaHQ/sigma.git"
SIGMA_REPO_DIR = "C:/Users/admin/Desktop/Python-For-Security-Engineer/lesson_10/sigma"
SIGMA_RULES_DIR = os.path.join(SIGMA_REPO_DIR, "rules", "windows", "raw_access_thread")
OUTPUT_FILE = "C:/Users/admin/Desktop/Python-For-Security-Engineer/lesson_10/sigma_rule.txt"
EXCEL_FILE = "C:/Users/admin/Desktop/Python-For-Security-Engineer/lesson_10/output.xlsx"

def pull_sigma_rules():
    if os.path.exists(SIGMA_REPO_DIR):
        subprocess.run(["git", "-C", SIGMA_REPO_DIR, "pull"], check=True)
    else:
        subprocess.run(["git", "clone", SIGMA_REPO_URL, SIGMA_REPO_DIR], check=True)

def convert_sigma_to_splunk():
    with open(OUTPUT_FILE, 'w') as f:
        for root, _, files in os.walk(SIGMA_RULES_DIR):
            for file in files:
                try:
                    if file.endswith(".yml") or file.endswith(".yaml"):
                        sigma_file_path = os.path.join(root, file)
                        subprocess.run(["sigma", "convert", "--without-pipeline", "-t", "splunk", sigma_file_path], stdout=f, check=True)
                except Exception as e:
                    print("Error: ", e)

def write_to_excel():
    with open(OUTPUT_FILE, 'r') as f:
        data = [{"rule": line.strip(), "description": ""} for line in f]
    df = pd.DataFrame(data)
    df.to_excel(EXCEL_FILE, index=False)

def job():
    pull_sigma_rules()
    convert_sigma_to_splunk()
    write_to_excel()

def main():
    schedule.every(3).seconds.do(job)
    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == "__main__":
    main()
