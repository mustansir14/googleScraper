import undetected_chromedriver as uc
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import gspread
import logging, os, random, time
from datetime import datetime
logging.basicConfig(format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %H:%M:%S', level=logging.INFO)
os.environ['WDM_LOG'] = str(logging.NOTSET)

def excel_style(row, col):
    """ Convert given row and column number to an Excel-style cell name. """
    LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    result = []
    while col:
        col, rem = divmod(col-1, 26)
        result[:0] = LETTERS[rem]
    return ''.join(result) + str(row)

SHEETS = {
    "United States": "1Cxp3hA6wu-fOJ1oynT3reFPiBowOjDlLOVm5tC_o5Os",
    "Taiwan": "1Ce6DBGVsXM1MnhcgbyS2wIQcjy8IqN4grjIcb8495vU",
    "Australia": "1QOGt3A4QN1krq8Kpfgj2Zchu9W9QkypY4aTYlQFnS7w",
    "Spain": "1-2ncwONVf9HC-GiVCAHjvCioaelbFwkcrHtSg4qkC10",
    "Netherlands": "1IjX88Vql0duEMP6sTZsEu3G-PT9OPY_Y4jHChnVLvtc",
    "Japan": "1TOHwSRNrKSuCV6CP2mtW3GY4CqM4LOqCkHkI5p7p0xE",
    "Italy": "1MGs2u62CZH7cscAvBS74DpeCuhKsHRYFAaLWIYhosdc",
    "Hong Kong": "16lSO6PjdAdXt76GgwDMef3gh85W9IYKy-AOI1g9rMR0",
}

def get_links(driver, url):

    unique_links = []
    count = 0
    while len(unique_links) < 20:
        if len(unique_links) == 0:
            driver.get(url)
        else:
            driver.get(url + f"&start={count}")
        links = [x.find_element(By.XPATH, "..").get_attribute("href") for x in driver.find_elements(By.TAG_NAME, "h3")]
        if len(links) == 0:
            raise Exception("Found no links... Blocked")
        for link in links:
            if link and link.strip() and link not in unique_links:
                unique_links.append(link.split("#")[0])
            if len(unique_links) == 20:
                break
        count += 10
        time.sleep(random.randint(30, 60))
    return unique_links

if __name__ == "__main__":

    gc = gspread.service_account() 
    
    for country, key in SHEETS.items():
        logging.info("Country: " + country)
        sh = gc.open_by_key(key)
        input_sheet = sh.get_worksheet(0)
        queries = input_sheet.col_values(1)[1:]
        urls = input_sheet.col_values(2)[1:]
        for query, url in zip(queries, urls):
            logging.info("Searching query: " + query)
            options = Options()
            options.headless = True
            options.add_argument("window-size=1920,1080")
            options.add_argument("--no-sandbox")
            options.add_argument('--disable-dev-shm-usage') 
            driver = uc.Chrome(driver_executable_path=ChromeDriverManager().install(), options=options)
            try:
                work_sheet = sh.worksheet(query)
            except gspread.exceptions.WorksheetNotFound:
                try:
                    work_sheet = sh.add_worksheet(title=query, rows=100, cols=100)
                except Exception as e:
                    logging.error(str(e))
                    continue
            work_sheet.update("A1:B1", [["Query:", query]])
            work_sheet.format('A1:B1', {'textFormat': {'bold': True}})
            work_sheet.update("A3:A23", [["Position"]] + [[x] for x in range(1, 21)])
            dates = work_sheet.row_values(3)[1:]
            today_date = datetime.now().strftime("%d-%m-%Y")
            if today_date in dates:
                col_num = dates.index(today_date) + 2
            else:
                col_num = len(dates) + 2
                work_sheet.update_cell(3, col_num, today_date)
            try:
                links = get_links(driver, url)
            except Exception as e:
                logging.error(str(e))
                continue
            try:
                driver.quit()
            except:
                pass
            logging.info("Found results!!")
            work_sheet.update(f'{excel_style(4, col_num)}:{excel_style(23, col_num)}', [[x] for x in links])
            logging.info(str(len(links)) + " links added for query: " + query)

            # pivot table
            work_sheet.update("A28:B28", [["URL", "Positions"]])
            dates = work_sheet.row_values(3)[1:]
            work_sheet.update(f'B29:{excel_style(29, 1+len(dates))}', [dates])
            url_rows = work_sheet.get_all_values()[3:23]
            unique_urls = []
            for row in url_rows:
                for url in row[1:]:
                    if url.strip() and url not in unique_urls:
                        unique_urls.append(url)
            work_sheet.update(f'A30:A{30+len(unique_urls)}', [[x] for x in unique_urls])
            positions = []
            for unique_url in unique_urls:
                positions.append([])
                for i in range(1, len(url_rows[0])):
                    found = False
                    for j in range(20):
                        if url_rows[j][i] == unique_url:
                            positions[-1].append(j+1)
                            found = True
                            break
                    if not found:
                        positions[-1].append('-')
                    
            work_sheet.update(f'B30:{excel_style(30+len(unique_urls), 1+len(dates))}', positions)
            work_sheet.format(f'B30:{excel_style(30+len(unique_urls), 1+len(dates))}', {"horizontalAlignment": "RIGHT"})
            time.sleep(random.randint(30, 60))

        

        
        
        



        

