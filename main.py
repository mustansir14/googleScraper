import undetected_chromedriver as uc
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import gspread
import logging, os
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
    "United States": "1Cxp3hA6wu-fOJ1oynT3reFPiBowOjDlLOVm5tC_o5Os"
}

def get_links(driver, url):

    driver.get(url)
    links = [x.find_element(By.XPATH, "..").get_attribute("href") for x in driver.find_elements(By.TAG_NAME, "h3")]
    unique_links = []
    for link in links:
        if link and link.strip() and link not in unique_links:
            unique_links.append(link.split("#")[0])
        if len(unique_links) == 20:
            break
    if len(unique_links) < 20:
        driver.get(url + "&start=10")
        links = [x.find_element(By.XPATH, "..").get_attribute("href") for x in driver.find_elements(By.TAG_NAME, "h3")]
        for link in links:
            if link and link.strip() and link not in unique_links:
                unique_links.append(link.split("#")[0])
            if len(unique_links) == 20:
                break
    return unique_links

if __name__ == "__main__":

    gc = gspread.service_account()
    options = Options()
    options.headless = True
    options.add_argument("window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument('--disable-dev-shm-usage')  
    driver = uc.Chrome(driver_executable_path=ChromeDriverManager().install(), options=options)
    
    sh = gc.open_by_key("1Cxp3hA6wu-fOJ1oynT3reFPiBowOjDlLOVm5tC_o5Os")
    input_sheet = sh.get_worksheet(0)
    queries = input_sheet.col_values(1)[1:]
    urls = input_sheet.col_values(2)[1:]
    for query, url in zip(queries, urls):
        logging.info("Searching query: " + query)
        try:
            work_sheet = sh.worksheet(query)
        except gspread.exceptions.WorksheetNotFound:
            work_sheet = sh.add_worksheet(title=query, rows=100, cols=100)
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
                if url not in unique_urls:
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

        

        
        
        



        
