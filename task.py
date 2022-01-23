import os
import time
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.FileSystem import FileSystem
from RPA.PDF import PDF

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
EXCEL_FILE_NAME = 'agencies.xlsx'
SHEET_AGENCIES_NAME = 'Agencies'
URL = 'https://itdashboard.gov/'
STD_TIMEOUT = 180
UII_URL_NAME = 'UII_URL'
EMPTY_UII_URL = '--'
NAME_INVESTMENT_LOCATOR = '1. Name of this Investment: '
UII_LOCATOR = '2. Unique Investment Identifier (UII): '

browser = Selenium()
excel = Files()
tables = Tables()
lib = FileSystem()
pdf = PDF()


def initial_configuration():
    """ Configurate the minimal configuration
    to run the processes.
    """
    create_directory(OUTPUT_DIRECTORY)
    open_or_create_excel_file(os.path.join(OUTPUT_DIRECTORY, EXCEL_FILE_NAME))


def get_list_of_agencies_and_save_in_excel():
    """
    Get the list of al agencies in https://itdashboard.gov/' and
    save the data in a sheet named Agencies of an excel file.
    """

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

    locator = 'xpath://a[@href="#home-dive-in"]'
    browser.wait_until_page_contains_element(locator, STD_TIMEOUT )
    browser.click_element(locator)


def get_agencies():
    locator = 'xpath://div[@id="agency-tiles-widget"]/div/div[contains(@class, "row")]'
    browser.wait_until_page_contains_element(locator + '[1]')
    agencies_blocks = browser.get_webelements(locator)
    agencies_list = []
    # Agencies is ordered by blocks (each block has three angencies)
    for agencies_block in agencies_blocks:
        agencies = agencies_block.find_elements_by_xpath('.//div[@class="tuck-5"]')
        for agency in agencies:
            url = agency.find_element_by_xpath('.//a').get_attribute('href')
            name = agency.find_element_by_xpath('.//span[contains(@class, "h4") and contains(@class, "w200")]').text
            amount = agency.find_element_by_xpath('.//span[contains(@class, "h1") and contains(@class, "w900")]').text
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
    
    table_element_agency, table_header_element_agency = get_table_element_from_url()
    table_agency = read_table_from_element(table_element_agency, table_header_element_agency)

    close_website()
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

def get_table_element_from_url():
    """
    Get the table elements with all "Individual Investments". 

    :return: WebElements of tables with data and header.
    :rtype: (WebElement, WebElement).
    """
    selection_locator = f'xpath://select[@name="investments-table-object_length"]'
    table_header_locator = f'xpath://div[@id="investments-table-object_wrapper"]//div[@class="dataTables_scrollHead"]//table'
    table_locator = f'xpath://table[@id="investments-table-object"]'
    paginate_locator_2 = f'xpath://div[@id="investments-table-object_paginate"]//a[not(contains(@class, "disabled")) and @data-dt-idx="3"]'
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
    # wait until table is available
    browser.wait_until_page_contains_element(table_locator, STD_TIMEOUT)
    table_element = browser.get_webelement(table_locator)
    return table_element, table_header_element

def read_table_from_element(table_element, table_header_element):
    """Parses and returns the given element tables as a Table structured.

    :param WebElement table_element: Table web element.
    :param WebElement table_header_element: Header of table web element.
    :return: Table structured.
    :rtype: Table.
    """
    table_header = []
    table_rows = []

    # Get table header and include it in a list.
    th_header_locator = './/tr[@role="row"]/th'
    th_header_elements = table_header_element.find_elements_by_xpath(th_header_locator)
    for index, th_element in enumerate(th_header_elements):
        table_header.append(th_element.text)
        if index == 0:
            table_header.append(UII_URL_NAME)
    
    # Get table data and include it in a list.
    tr_table_locator = './/tbody/tr[@role="row"]'
    tr_table_elements = table_element.find_elements_by_xpath(tr_table_locator)
    for tr_index, tr_element in enumerate(tr_table_elements):
        td_locator = './td'
        td_elements = tr_element.find_elements_by_xpath(td_locator)
        td_values = []
        for td_index, td_element in enumerate(td_elements):
            td_values.append(td_element.text)
            if td_index == 0:
                try:
                    td_values.append(td_element.find_element_by_xpath('./a').get_attribute("href"))
                except:
                    td_values.append(EMPTY_UII_URL)
        table_rows.append(td_values)

    # Create a table from lists
    return tables.create_table(data=table_rows, columns=table_header)

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
    browser.set_download_directory(directory=get_absolute_path_directory(directory))
    for url in urls:
        open_website(url)
        # Wait until page contains the button to download the pdf.
        browser.wait_until_page_contains_element(
            'xpath://div[@id="business-case-pdf"]/a', STD_TIMEOUT)
            
        browser.click_element('xpath://div[@id="business-case-pdf"]/a')
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
    return os.path.join(os.getcwd(), relative_path_directory)


def wait_until_download_end(directory, timeout):
    """When a pdf is downloading, checks that is not a temporal
    file .download (firefox, opera) or .crdownload (chrome).

    :param str directory: Path with PDFs files.
    :param int timeout: Time out in seconds.
    """
    seconds = 0
    while seconds < timeout:
        time.sleep(1)
        if not lib.find_files(os.path.join(directory,"*.*download")):
            break
        seconds += 1


def extract_data_from_pdf():
    """Extract "Name of this Investment" and "Unique Investment Identifier (UII)"
    and compare this values with the columns "Investment Title" and "UII" in Excel,
    and save the comparison in the Excel file.
    """
    title_and_uii_list = get_title_and_uii_list()
    pdf_list = get_pdf_list()
    pdf_name_and_uii_list = []
    for pdf_file in pdf_list:
        pdf_text = get_page_1_pdf_text(pdf_file)
        pdf_name_and_uii = get_pdf_name_and_uii(pdf_text)
        pdf_name_and_uii_list.append(pdf_name_and_uii)
    title_and_uii_comparison = compare_pdf_and_excel_title_and_uii(
        pdf_name_and_uii_list, title_and_uii_list)
    save_table_in_excel(title_and_uii_comparison, 'Title and UII Comparison')


def get_title_and_uii_list():
    """Get a list with columns title and uii of business cases from the Excel file.
    Only get the business cases with link, and a pdf downloaded.

    :return: Business cases list of dict with title and uii keys.
    :rtype: List of dict.
    """
    excel.read_worksheet(AGENCY_NAME.capitalize())
    data = excel.read_worksheet_as_table(header=True)
    tables.filter_table_by_column(data, UII_URL_NAME, '!=', EMPTY_UII_URL)
    title_and_uii_list = [{'title': title, 'uii': uii} for title, uii in zip(
        tables.get_table_column(data, 'Investment Title'), tables.get_table_column(data, 'UII'))]
    return title_and_uii_list


def get_pdf_list():
    """Get a list of the PDFs stored in the output directory ordered by date.

    :return: List with PDFs Files.
    :rtype: list of File.
    """

    pdf_list = lib.find_files(os.path.join(OUTPUT_DIRECTORY, "*.pdf"))
    pdf_list.sort(key=lambda x: x.mtime, reverse=False)
    return pdf_list


def get_page_1_pdf_text(pdf_file):
    """Get the text inside the first page of a pdf file.

    :param File pdf: Name of PDF file.
    :return: Text of PDF file.
    :rtype: str.
    """
    pdf_text = pdf.get_text_from_pdf(pdf_file.path)
    return pdf_text[1]


def get_pdf_name_and_uii(pdf_text):
    """Extract the Name and UII from a PDF text.

    :param str pdf_text: PDF text.
    :return: Dictionary with the PDF name and UII
    :rtype: Dict.
    """
    section_a_index = pdf_text.find('Section A:')
    section_b_index = pdf_text.find('Section B:')
    section_a_text = pdf_text[section_a_index:section_b_index]
    name_start = section_a_text.find(NAME_INVESTMENT_LOCATOR) + len(NAME_INVESTMENT_LOCATOR)
    name_end = section_a_text.find(UII_LOCATOR)
    uii_right_index = section_a_text.find(UII_LOCATOR) + len(UII_LOCATOR)
    name_text = section_a_text[name_start:name_end]
    uii_text = section_a_text[uii_right_index:]
    pdf_name_and_uii = {'pdf_name': name_text, 'pdf_uii': uii_text}
    return pdf_name_and_uii


def compare_pdf_and_excel_title_and_uii(pdf_name_and_uii_list, title_and_uii_list):
    """Compare the name and uii keys of the pdf file list with the title and uii keys 
    of the excel list, and insert another key with the comparison.  

    :param list pdf_name_and_uii_list: List of directories with PDFs names and UIIs keys.
    :param list title_and_uii_list: List with titles and UIIs keys extracted from agency sheet.
    :return: Tables with 
    :rtype: Table.
    """
    title_and_uii_comparison = [{**excel_item, **pdf_item, 'comparison title': excel_item['title'] == pdf_item['pdf_name'],
                                 'comparison uii': excel_item['uii'] == pdf_item['pdf_uii']} for excel_item, pdf_item in zip(title_and_uii_list, pdf_name_and_uii_list)]
    title_and_uii_comparison_table = tables.create_table(
        title_and_uii_comparison)
    return title_and_uii_comparison_table


# Define a main() function that calls the other functions in order:
def main():
    """
    Define a main() function that calls the other functions in order:
    """
    try:
        initial_configuration()
        get_list_of_agencies_and_save_in_excel()
        get_agency_investments_and_save_in_excel()
        download_pdf_with_agency_business_case()
        extract_data_from_pdf()

    finally:
        browser.close_all_browsers()
        excel.close_workbook()
        pdf.close_all_pdfs()


# Call the main() function, checking that we are running as a stand-alone script:
if __name__ == "__main__":
    main()
