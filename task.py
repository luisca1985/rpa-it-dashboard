import os
import time
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.FileSystem import FileSystem
from bs4 import BeautifulSoup

# Get the name of the agency selected.
try:
    AGENCY_NAME = os.environ["AGENCY_NAME"]
except:
    AGENCY_NAME = 'National Science Foundation'
    print(f'Problems with AGENCY_NAME, its values is: { AGENCY_NAME }')
else:
    print(
        f'AGENCY_NAME was configured successfully, its values is: { AGENCY_NAME }')

OUTPUT_DIRECTORY = 'output'
EXCEL_PATH = OUTPUT_DIRECTORY + '/agencies.xlsx'
SHEET_AGENCIES_NAME = 'Agencies'
URL = 'https://itdashboard.gov/'
STD_TIMEOUT = 180
UII_URL_NAME = 'UII_URL'
EMPTY_UII_URL = '--'

browser = Selenium()
excel = Files()
tables = Tables()
lib = FileSystem()


def get_list_of_agencies_and_save_in_excel():
    """
    Get the list of al agencies in https://itdashboard.gov/' and
    save the data in a sheet named Agencies of an excel file.
    """
    create_directory(OUTPUT_DIRECTORY)
    open_or_create_excel_file(EXCEL_PATH)
    open_website(URL)
    click_div_in()
    agencies = get_agencies()
    close_website()
    save_table_in_excel(agencies, sheet_name=SHEET_AGENCIES_NAME)


def create_directory(dir_name):
    """
    Create the directory.

    :param str dir_name: The name of the directory.
    """
    if lib.does_directory_exist(dir_name):
        print(f'The directory { dir_name } already exist.')
    else:
        print(f'The directory { dir_name } is created.')
    lib.create_directory(dir_name)


def open_or_create_excel_file(excel_path):
    """
    Open the excel file to save the spend amounts for each agency and 
    individual investments of selected agency.

    :param str excel_path: The name of excel file and path where it is located
    """
    try:
        excel.open_workbook(excel_path)
    except:
        excel.create_workbook(excel_path)
        print(f'The Excel file { excel_path } is created.')
    else:
        print(
            f'The Excel file { excel_path } already exist, and it is opened.')


def open_website(url):
    """
    Open a website using the browser available.

    :param str url: Website URL.
    """
    browser.open_available_browser(url)


def close_website():
    """
    Close the active browser.
    """
    browser.close_browser()


def click_div_in():
    """
    Click "DIVE IN" on the homepage to reveal the spend amounts for each agency.
    """
    locator = 'xpath://*[@id="node-23"]/div/div/div/div/div/div/div/a'
    browser.wait_until_page_contains_element(locator)
    browser.click_element(locator)


def get_agencies():
    locator = 'xpath://*[@id="agency-tiles-widget"]/div/div'
    browser.wait_until_page_contains_element(locator + '[1]')
    agencies_blocks = browser.get_webelements(locator)
    agencies_list = []
    # Agencies is ordered by blocks (each block has three angencies)
    for agencies_block in agencies_blocks:
        agencies = agencies_block.find_elements_by_xpath('./div')
        for agency in agencies:
            info = agency.find_element_by_xpath('./div/div/div/div[1]/a')
            url = info.get_attribute('href')
            name = info.find_element_by_xpath('./span[1]').text
            amount = info.find_element_by_xpath('./span[2]').text
            agency_dict = {
                "name": name.capitalize(), "amount": amount, "url": url}
            agencies_list.append(agency_dict)
    return tables.create_table(agencies_list)


def save_table_in_excel(table, sheet_name):
    """
    Save a table in a sheet in the active excel workbook.

    :param Table table: The table with the data.
    :param str sheet_name: The name of the sheet with the data.
    """
    try:
        excel.remove_worksheet(name=sheet_name)
    except:
        print(f"Sheet { sheet_name } don't exist.")
    else:
        print(f"Sheet { sheet_name } already exists, so it is removed.")
    finally:
        excel.create_worksheet(name=sheet_name, content=table, header=True)
        print(f"A new sheet { sheet_name } is created.")
        excel.save_workbook()


def get_agency_investments_and_save_in_excel():
    """
    Get table with all Individual Investments of the selected agency and
    write it to a new sheet in excel.
    """
    agency = extract_table_and_filter(
        SHEET_AGENCIES_NAME, 'name', '==', AGENCY_NAME.capitalize())
    urls_agency = get_urls_from_table(agency, 'url')
    open_website(urls_agency[0])
    html_table_agency, html_table_header_agency = get_table_html_from_url()
    close_website()
    table_agency = read_table_from_html(
        html_table_agency, html_table_header_agency)
    save_table_in_excel(table_agency, sheet_name=AGENCY_NAME.capitalize())


def extract_table_and_filter(sheet_name, column, operator, value):
    """
    Extract a table from a sheet of excel and filter it using the column, 
    the logical operator and the values to compare.

    :param str sheet_name: The name of the sheet with de table.
    :param str column: Column with the data used to filter.
    :param str operator: Logical operator used to filter the table.
    :param str value: Value which is compare with the column values.
    :type priority: integer or None
    :return: the data filter
    :rtype: Table
    """
    excel.read_worksheet(sheet_name)
    data = excel.read_worksheet_as_table(header=True)
    tables.filter_table_by_column(data, column, operator, value)
    return data


def get_urls_from_table(table, url_column_name='url'):
    """
    Gets urls in a column.

    :param Table table: Table with url column.
    :param str url_column_name: Column with the urls.
    :return: List with urls.
    :rtype: List.
    """
    return tables.get_table_column(table, url_column_name)


def get_table_html_from_url():
    """
    Get a HTML table with all "Individual Investments". 

    :return: HTML source of tables with data and header.
    :rtype: str.
    """
    table_id = 'investments-table-object'
    selection_locator = f'css:#{ table_id }_length > label > select'
    table_header_locator = f'css:#{ table_id }_wrapper > div.dataTables_scroll > div.dataTables_scrollHead > div > table'
    table_locator = f'css:#{ table_id }'
    paginate_locator_2 = f'xpath://*[@id="{ table_id }_paginate"]/span/a[2]'
    # wait until selection component is available
    browser.wait_until_page_contains_element(selection_locator, STD_TIMEOUT)
    # wait until paginate 2 is available
    browser.wait_until_page_contains_element(paginate_locator_2, STD_TIMEOUT)
    # select to see all table rows
    browser.select_from_list_by_value(selection_locator, "-1")
    # wait until paginate 2 is not available
    browser.wait_until_page_does_not_contain_element(
        paginate_locator_2, STD_TIMEOUT)
    # wait until table header is available
    browser.wait_until_page_contains_element(table_header_locator, STD_TIMEOUT)
    table_header_element = browser.get_webelement(table_header_locator)
    table_header_html = table_header_element.get_attribute('innerHTML')
    # wait until table is available
    browser.wait_until_page_contains_element(table_locator, STD_TIMEOUT)
    table_element = browser.get_webelement(table_locator)
    table_html = table_element.get_attribute('innerHTML')
    # print(table_header_html)
    return table_html, table_header_html


def read_table_from_html(html_table, html_table_header):
    """Parses and returns the given HTML tables as a Table structured.

    :param str html_table: Table HTML markup.
    :param str html_table_header: Header of table HTML markup.
    :return: Table structured.
    :rtype: Table.
    """
    table_header_rows = []
    table_rows = []
    table_header = BeautifulSoup(html_table_header, "html.parser")
    table = BeautifulSoup(html_table, "html.parser")

    # Get table header and include it in a list.
    for table_row in table_header.select('tr'):
        cells = table_row.find_all('th')

        if len(cells) > 0:
            cell_values = []

            for index, cell in enumerate(cells):
                cell_values.append(cell.text.strip())
                if index == 0:
                    cell_values.append(UII_URL_NAME)

            table_header_rows.append(cell_values)

    # Get table data and include it in a list.
    for table_row in table.select('tr'):
        cells = table_row.find_all('td')

        if len(cells) > 0:
            cell_values = []

            for index, cell in enumerate(cells):
                cell_values.append(cell.text.strip())
                if index == 0:
                    try:
                        cell_values.append(URL + cell.find('a')['href'])
                    except:
                        cell_values.append(EMPTY_UII_URL)

            table_rows.append(cell_values)

    # Delete unnecessary rows
    table_rows.pop(0)

    # Create a table from lists
    return tables.create_table(data=table_rows, columns=table_header_rows[0])


def download_pdf_with_agency_business_case():
    """If the "UII" column contains a link, open it and download PDF with 
    Business Case.
    """
    table_uii_with_urls = extract_table_and_filter(
        AGENCY_NAME.capitalize(), UII_URL_NAME, '!=', EMPTY_UII_URL)
    urls_uii = get_urls_from_table(table_uii_with_urls, UII_URL_NAME)
    download_documents_from_urls(urls_uii)


def download_documents_from_urls(urls):
    """Download PDFs with Business Case from a list of urls.

    :param str urls: Urls get page where PDFs with Business Case can be downloaded.
    """
    # Create the directory that contents the PDFs using as a name, the agency name and the datetime.
    directory = OUTPUT_DIRECTORY
    browser.set_download_directory(
        directory=get_absolute_path_directory(directory))
    for url in urls:
        open_website(url)
        # Wait until page contains the button to download the pdf.
        browser.wait_until_page_contains_element(
            'css:#business-case-pdf > a', STD_TIMEOUT)
        browser.click_element('css:#business-case-pdf > a')
        # Wait until pdf starts to download.
        lib.wait_until_modified(directory, STD_TIMEOUT)
        # Wait until pdf finishes to download.
        wait_until_download_end(directory, STD_TIMEOUT)
        close_website()


def get_absolute_path_directory(relative_path_directory):
    """Transform the relative path to an absolute path.

    :param str relative_path_directory: Relative path.
    :return: Absolute path.
    :rtype: str.
    """
    return os.getcwd() + '/' + relative_path_directory


def wait_until_download_end(directory, timeout):
    """When a pdf is downloading, checks that is not a temporal
    file .download (firefox, opera) or .crdownload (chrome).

    :param str directory: Path with PDFs files.
    :param int timeout: Time out in seconds.
    """
    seconds = 0
    while seconds < timeout:
        time.sleep(1)
        if not lib.find_files(directory + "/*.*download"):
            break
        seconds += 1


# Define a main() function that calls the other functions in order:
def main():
    """
    Define a main() function that calls the other functions in order:
    """
    try:
        get_list_of_agencies_and_save_in_excel()
        get_agency_investments_and_save_in_excel()
        download_pdf_with_agency_business_case()

    finally:
        browser.close_all_browsers()
        excel.close_workbook()


# Call the main() function, checking that we are running as a stand-alone script:
if __name__ == "__main__":
    main()
