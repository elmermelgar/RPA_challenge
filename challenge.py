# +
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem
from RPA.Desktop.OperatingSystem import OperatingSystem
import time

browser_lib = Selenium()
excel_lib = Files()
file_lib= FileSystem()
ops = OperatingSystem()

##Variables 
dive_in = "//a[contains(@href,'home-dive-in')]"
span_list = "//a[contains(@href,'drupal/summary')][span]"
individual_agency_xpath="//a[contains(@href,'drupal/summary')]/span[text()='"
table_agency_headers_xpath= '//*[@class="dataTables_scrollHeadInner"]/table/thead/tr[2]/th'
all_rows_xpath='//*[@id="investments-table-object"]/tbody/tr'
text_file_content="National Science Foundation"
individual_investments_row_items_link_xpath= '//*[@id="investments-table-object"]/tbody/tr/td/a'
pdf_business_case_xpath='//*[@id="business-case-pdf"]/a'
text_file_path = "output/challenge_text.txt"
excel_file_path="output/challenge_book.xlsx"
select_table_xpath="//select[@name='investments-table-object_length']/option[text()='All']"
foot_table_xpath="//a[@data-dt-idx='3'][text()='2']"

def open_the_website(url):
    browser_lib.open_available_browser(url)

##find the DAVE IN button and press click
def search_element_an_click(xpath):
    browser_lib.wait_until_page_contains_element(xpath,120000)
    browser_lib.scroll_element_into_view(xpath)
    browser_lib.click_element(xpath) 

##Create and save the list of agencies with their amounts
def create_list_of(table_xpath):
    print(["[Create list of Agencies] Process starting..."])
    table_items = browser_lib.find_elements(table_xpath)
    save_headers_to_excel('Agencies', 0)
    save_items_to_excel(table_items,'span')

##function that reads the content of a text file and looks for the element on the page
def read_file_for(file_path):
    print("[Read File] Starting process...")
    if(file_lib.does_directory_exist("output")==False):
        file_lib.create_directory("output")
        print("Folder Not Exist")
    if validate_exist_file(file_path)==False:
        create_file_for(file_path, text_file_content)
        print("[File_Not_Exist] File Created")
    content = file_lib.read_file(file_path, encoding="utf-8")
    agency_button_xpath=individual_agency_xpath+content+"'"+"]"
    navigate_to_agency(agency_button_xpath)

## check if the folder and file exist
def validate_exist_file(agency_file:str):
    print("[Validation] Exist file...")
    if(file_lib.does_directory_exist("output")==False):
        file_lib.create_directory("output")
    if file_lib.does_file_exist(agency_file):
        return True
    else:
        return False

## function used to navigate an agency
def navigate_to_agency(agency_xpath: str):
    print("[Navegate to Agency] Process starting...")
    browser_lib.click_element(agency_xpath)
    browser_lib.wait_until_page_contains_element(table_agency_headers_xpath, 120000)
    browser_lib.scroll_element_into_view(table_agency_headers_xpath)

##Create the file with the content to search
def create_file_for(agency_file:str, text_file):
    file_lib.create_file(agency_file, text_file, encoding='utf-8', overwrite=True)

##Save individual investments in an Excel table
def save_table_individual_investments():
    print("[Save table individual invesments] Starting process...")
    save_headers_to_excel('individual_investments', table_agency_headers_xpath)
    browser_lib.find_element(select_table_xpath).click()
    browser_lib.wait_until_page_does_not_contain_element(foot_table_xpath, 120000)
    browser_lib.wait_until_page_contains_element(all_rows_xpath, 120000)
    browser_lib.scroll_element_into_view(all_rows_xpath)
    all_rows=browser_lib.find_elements(all_rows_xpath)
    save_items_to_excel(all_rows, 'td')
    
##Download the pdf files of a specific agency
def download_pdf_agency():
    print("[Download pdf agency] Starting process...")
    all_links = browser_lib.find_elements(individual_investments_row_items_link_xpath)
    list_of_links=[]
    for alllink in all_links:
        list_of_links.append(alllink.get_attribute('href'))
    for i in range(0, len(list_of_links)):
        try:  
            browser_lib.go_to(list_of_links[i])
            validate_if_element_exists(pdf_business_case_xpath)
            time.sleep(10)
        except Exception as e:
            print(e)
            pass     
    time.sleep(8)
    path_download="C:/Users/"+ops.get_username()+"/Downloads"
    matches = file_lib.find_files(path_download+"/*.pdf")
    file_lib.move_files(matches, "output",overwrite=True)

##Verify that the pdf file link is available
def validate_if_element_exists(xpath):
    tries = 0
    while True:
        if tries >= 10: 
            break
        try:
            browser_lib.set_download_directory("output/",download_pdf=True)
            browser_lib.wait_until_page_contains_element(xpath, 120000)
            browser_lib.scroll_element_into_view(xpath)
            browser_lib.click_element(xpath)
            break
        except Exception as e:
            tries += 1
            time.sleep(5)
            print(f"Error on {e}")

##Save the table values to an Excel file
def save_items_to_excel(table_items, tag_name):
    row=2
    for table_item in table_items:
        column=1
        elements = table_item.find_elements_by_tag_name(tag_name)
        for element in elements:
            excel_lib.set_cell_value(row=row, column=column, value=element.text)
            column=column+1
        row=row+1
    excel_lib.save_workbook()
    excel_lib.close_workbook()

##Create or open an Excel file and save the table headers
def save_headers_to_excel(sheet_name, table_headers_xpath):
    if validate_exist_file(excel_file_path)==False:
        excel_lib.create_workbook(excel_file_path, fmt="xlsx")
        excel_lib.rename_worksheet('Sheet', sheet_name)
        print("[Excel file not exist] File created...")
    else:
        excel_lib.open_workbook(excel_file_path)
        if (excel_lib.worksheet_exists(sheet_name)== False):
            excel_lib.create_worksheet(sheet_name, content=None, exist_ok=False, header=False)
        else:
            excel_lib.remove_worksheet(sheet_name)
            excel_lib.create_worksheet(sheet_name, content=None, exist_ok=False, header=False)
    excel_lib.set_active_worksheet(sheet_name)
    if(table_headers_xpath==0):
        excel_lib.set_cell_value(row=1, column=1, value="Agencies")
        excel_lib.set_cell_value(row=1, column=2, value="Amounts")
    else:
        table_headers = browser_lib.find_elements(table_headers_xpath)
        col=1
        for table_header in table_headers:
            excel_lib.set_cell_value(row=1, column=col, value=table_header.text)
            col=col+1

##Main system function that calls all other functions
def main():
    try:
        open_the_website("https://itdashboard.gov/")
        search_element_an_click(dive_in)
        create_list_of(span_list)
        read_file_for(text_file_path)
        save_table_individual_investments()
        download_pdf_agency()

    finally:
        browser_lib.close_all_browsers()
        pass


# Call the main() function, checking that we are running as a stand-alone script:
if __name__ == "__main__":
    main()
