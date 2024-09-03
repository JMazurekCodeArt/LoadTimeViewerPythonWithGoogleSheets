"""
A tool for recording the load times of websites stored in a Google Sheets document.

CREATING A GOOGLE SHEET
- Create a project in Google Cloud
- APIs & Services >>> Enable APIs & Services >>> Google Drive API >>> Enable and do the same for Google Sheets API
- Credentials >>> Create Credentials >>> Service Account >>> Fill in the details >>> Create and Continue
- Add a role: Basic > Editor >>> Done (Later, you can set more permissions for other users)
- Select the created service account >>> Keys >>> Add Key >>> Create New Key >>> JSON
- A JSON file will be downloaded. Open it and copy the contents to a file named `credentials.json`.
- Create a New Sheet >>> Name it >>> Share >>> Add the email address from the `credentials.json` file (line 6) as an editor
- Check in Google Sheets File >>> Settings >>> Locale Settings >>> "UNITED STATES"

EDITING DATA
- At the beginning of the program, there are variables that need to be changed for the program to work correctly.

PROGRAM DESCRIPTION AND FUNCTIONALITY
The program is based on the Selenium tool, which allows analyzing page load times. In the `load_site()` function, the program retrieves:
- `navigationStart` - the moment the request is sent,
- `responseStart` - the moment the browser receives the first byte of data,
- `domComplete` - the moment the entire page is loaded.

The URLs are taken from the sheet names. The URLs must be complete (https://address.com/).
Just add a new sheet, and the rest will be done automatically.

Each page is loaded three times, and the median is calculated.
After each load, the cache is cleared to ensure reliable results.

The data is recorded using the `save()` function as (Date, Time, Address, Backend (s), Frontend (s), Total Time (s)).

Additionally, thanks to the `ensure_headers` function, with each update, the average times of the last updated day are displayed.
The program also clears records. To edit the number of days for data retention, you need to change the `days` variable.
"""

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from functools import lru_cache
from datetime import datetime, timedelta
import pygsheets
import requests
import os

""" DATA TO BE CHANGED """
ark = "Website Load Time Monitoring"   # Name of the entire spreadsheet in Google Sheets
naz = "Available"                      # Name of the sheet to be edited, which will be skipped during website checks. It is intended for analyzing the other sheets
days = 30                              # Number of days for which data should be retained in Google Sheets


CREDENTIALS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials.json")


def get_urls_from_sheets():     # Main function that retrieves and checks URLs and calls subsequent functions.
    client = pygsheets.authorize(service_account_file=CREDENTIALS_FILE)
    spreadsheet = client.open(ark)
    for worksheet in spreadsheet.worksheets():
        url = worksheet.title
        try:
            response = requests.head(url, allow_redirects=True, timeout=30)
            if response.status_code == 200:
                print(f"Processing sheet with URL: {url}")
                avg_time(url)
                ensure_headers(url)
                clean_old_records(url)
        except requests.RequestException:
            print("Pominięto arkusz: ", url)
            if url != naz:
                date = datetime.now()
                save(date, url, "Błąd", "Brak", "połączenia")
                ensure_headers(url)
                clean_old_records(url)


def avg_time(url):      # Function that selects the median. You can edit the number of tests per site here.
    backend_times = []
    frontend_times = []
    whole_times = []
    print(' ')
    date = datetime.now()
    print('Date: ', date)
    print('Test for ', url)
    for x in range(1, 4):
        print('Test', x)
        backend_times.append(load_site(url)[0])
        frontend_times.append(load_site(url)[1])
        whole_times.append(load_site(url)[2])
        load_site.cache_clear()
        print('--------------')

    zipped = list(zip(whole_times, backend_times, frontend_times))

    zipped.sort(key=lambda x: x[0])

    whole_times, backend_times, frontend_times = zip(*zipped)

    middle_backend_time = backend_times[1]
    middle_frontend_time = frontend_times[1]
    middle_whole_time = whole_times[1]

    avg_backend_time = round(middle_backend_time, 3)
    avg_frontend_time = round(middle_frontend_time, 3)
    avg_whole_time = round(middle_whole_time, 3)
    print('Average time is', avg_whole_time, 'seconds')
    save(date, url, avg_backend_time, avg_frontend_time, avg_whole_time)


@lru_cache(maxsize=128)
def load_site(url):     # Function that measures the load times of a webpage.
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(options=chrome_options)
    driver.get(url)

    navigation_start = driver.execute_script("return window.performance.timing.navigationStart")
    response_start = driver.execute_script("return window.performance.timing.responseStart")
    dom_complete = driver.execute_script("return window.performance.timing.domComplete")

    backend_performance_calc = (response_start - navigation_start) / 1000
    frontend_performance_calc = (dom_complete - response_start) / 1000
    performance_calc = (dom_complete - navigation_start) / 1000

    print("Back End: %s s" % backend_performance_calc)
    print("Front End: %s s" % frontend_performance_calc)
    print("All: %s s" % performance_calc)
    driver.quit()
    return backend_performance_calc, frontend_performance_calc, performance_calc


def save(date, url, avg_backend_time, avg_frontend_time, avg_whole_time):       # Function that saves records in the Sheet.
    client = pygsheets.authorize(service_account_file=CREDENTIALS_FILE)

    new_date = date.strftime("%d/%m/%Y")
    hour = date.strftime("%H:%M:%S")
    values = [new_date, hour, url, avg_backend_time, avg_frontend_time, avg_whole_time]

    spreadsht = client.open(ark)
    worksht = spreadsht.worksheet("title", url)

    worksht.insert_rows(4, number=1, values=values)


def ensure_headers(url):        # Filling in the headers and the top part of the sheet.
    client = pygsheets.authorize(service_account_file=CREDENTIALS_FILE)
    spreadsheet = client.open(ark)
    sheet = spreadsheet.worksheet_by_title(url)

    headers_list = [
        ["", "", "", "=MAXIFS(D5:D, A5:A, A5)", "=MAXIFS(E5:E, A5:A, A5)", "=MAXIFS(F5:F, A5:A, A5)", "Najdłuższy czas z dnia: ", "=A5"],
        ["", "", "", "=MINIFS(D5:D, A5:A, A5)", "=MINIFS(E5:E, A5:A, A5)", "=MINIFS(F5:F, A5:A, A5)", "Najkrótszy czas z dnia: ", "=A5"],
        ["", "", "", "=ŚREDNIA.WARUNKÓW(D5:D, A5:A, A5)", "=ŚREDNIA.WARUNKÓW(E5:E, A5:A, A5)", "=ŚREDNIA.WARUNKÓW(F5:F, A5:A, A5)", "Średni czas z dnia: ", "=A5"],
        ["Data", "Godzina", "Adres", "Backend (s)", "Frontend (s)", "Całkowity czas (s)"]
    ]

    for i, headers in enumerate(headers_list, start=1):
        existing_headers = sheet.get_row(i, include_tailing_empty=False)
        if existing_headers != headers:
            sheet.update_row(i, headers)


def clean_old_records(url):     # Cleaning up old records.
    client = pygsheets.authorize(service_account_file=CREDENTIALS_FILE)
    today = datetime.now()
    cutoff_date = today - timedelta(days=days)
    spreadsheet = client.open(ark)
    worksheet = spreadsheet.worksheet("title", url)
    rows = worksheet.get_all_values(include_tailing_empty=False, include_tailing_empty_rows=False)
    headers = rows[3]
    date_col_index = headers.index('Data')
    filtered_rows = [headers]

    for row in rows[4:]:
        try:
            record_date = datetime.strptime(row[date_col_index], "%d/%m/%Y")

            if record_date >= cutoff_date:
                filtered_rows.append(row)
        except (ValueError, TypeError, IndexError) as e:
            print(f"Error processing row: {row}, Error: {e}")
            continue

    if len(rows) > 4:
        worksheet.clear(start="A5")

    if len(filtered_rows) > 1:
        worksheet.update_values('A5', filtered_rows[1:])


if __name__ == "__main__":
    get_urls_from_sheets()
