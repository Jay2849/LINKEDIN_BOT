from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import utils

# Setup Chrome driver
options = Options()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Load message template and keywords
print("Loading message template...")
message_template = utils.load_message_template("SendText.txt")

print("Loading keywords...")
keywords = utils.load_keywords_from_excel("keywords.xlsx")
print(f"Loaded {len(keywords)} keywords.")

# Go through each keyword
for keyword in keywords:
    print(f"Searching jobs for keyword: {keyword}")
    utils.search_jobs(driver, keyword)
    jobs = utils.get_filtered_job_elements(driver)

    for job in jobs:
        details = utils.extract_job_details(driver, job)
        if not details:
            continue
        if utils.is_hr_related(details['description']):
            print("HR related job skipped.")
            continue

        message = utils.personalize_message(message_template, details)
        utils.send_message(driver, job, message)

print("Job automation completed.")
driver.quit()
