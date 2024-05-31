import subprocess
import os
import pandas as pd
import schedule
import time

def pull_sigma_rules(repo_url, repo_dir):
    if os.path.exists(repo_dir):
        subprocess.run(["git", "-C", repo_dir, "pull"], check=True)
    else:
        subprocess.run(["git", "clone", repo_url, repo_dir], check=True)

def convert_sigma_to_splunk(sigma_dir, output_file):
    sigmac_path = "C:/Users/admin/anaconda3/Scripts/sigmac.exe"  # Thay đổi nếu cần thiết
    if not os.path.isfile(sigmac_path):
        raise FileNotFoundError(f"{sigmac_path} không được tìm thấy")

    with open(output_file, 'w') as f:
        for root, dirs, files in os.walk(sigma_dir):
            for file in files:
                if file.endswith(".yml") or file.endswith(".yaml"):
                    sigma_file_path = os.path.join(root, file)
                    subprocess.run([sigmac_path, "convert", "-t", "splunk", sigma_file_path], stdout=f, check=True)

def write_to_excel(data, excel_file):
    df = pd.DataFrame(data)
    df.to_excel(excel_file, index=False)

def job():
    repo_url = "https://github.com/SigmaHQ/sigma.git"
    repo_dir = "C:/Users/admin/Desktop/Python-For-Security-Engineer/lesson_10/sigma"
    sigma_dir = os.path.join(repo_dir, "rules", "windows", "process_creation")
    output_file = "C:/Users/admin/Desktop/Python-For-Security-Engineer/lesson_10/sigma_rule.txt"
    excel_file = "C:/Users/admin/Desktop/Python-For-Security-Engineer/lesson_10/output.xlsx"

    pull_sigma_rules(repo_url, repo_dir)
    convert_sigma_to_splunk(sigma_dir, output_file)
    
    with open(output_file, 'r') as f:
        data = [{"rule": line.strip(), "description": ""} for line in f]
    write_to_excel(data, excel_file)

schedule.every(3).seconds.do(job)

while True:
    schedule.run_pending()
    time.sleep(1)
