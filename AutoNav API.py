import requests
import re
import openpyxl
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time

#  Substitute the chromedriver path with your actual path
CHROMEDRIVER_PATH = "Your/Path/To/chromedriver"

# For static loading page method（use requests library）
def fetch_with_requests(school_name):
    url = f"https://xxxx.xxxx.com/item/{school_name}"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        resp = requests.get(url, headers=headers, timeout=10)
        resp.encoding = resp.apparent_encoding
        html = resp.text.replace('\xa0', ' ').replace('&nbsp;', '')

        patterns = [
            r"Student quantities[:：]?\s*([\d\.thousands people]+)", # (Substitute with actual patterns)
            r"Bechelors,Masters,and Phds([\d\.thousands people]+)"
        ]
        for pattern in patterns:
            match = re.search(pattern, html)
            if match:
                return match.group(1).strip()
    except Exception as e:
        print(f" requests failed: {e}")
    return None

# For dynamically loading pages（used Selenium library）
def fetch_with_selenium(school_name):
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    service = Service(CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)

    try:
        url = f"Your url" # (Substitute with actual URL)
        driver.get(url)
        time.sleep(4)

        html = driver.page_source.replace('\xa0', ' ')
        infobox_matches = re.findall(r"(Students).*?<dd[^>]*>(.*?)</dd>", html, re.DOTALL) # (Substitute with actual pattern)
        if infobox_matches:
            return re.sub('<.*?>', '', infobox_matches[0][1]).strip()
    except Exception as e:
        print(f"use selenium  {school_name} failed: {e}")
    finally:
        driver.quit()
    return None

# Main function to fetch student counts for a list of schools

def fetch_all_student_counts(school_list):
    result = {}
    for school in school_list:
        print(f"\n Grabing {school} result...")
        count = fetch_with_requests(school)
        if not count:
            print("Requests failed, trying with selenium...")
            count = fetch_with_selenium(school)
        if not count:
            count = "No result"
        print(f"{school}: {count}")
        result[school] = count
    return result

# Save results to Excel
def save_to_excel(data: dict, filename="xxxx.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "xxxx" # (Substitute with actual title)
    ws.append(["xxxx", "xxxx"]) # (Substitute with actual headers)
    for school, count in data.items():
        ws.append([school, count])
    path = os.path.join(os.getcwd(), filename)
    wb.save(path)
    print(f"\n Excel Exported to：{path}")

# Example usage
if __name__ == "__main__":
    school_list = ["xxxx"] # (Substitute with actual school names)
    data = fetch_all_student_counts(school_list)
    save_to_excel(data)
