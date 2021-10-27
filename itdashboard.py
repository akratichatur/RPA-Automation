import time
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Tables import Tables
from RPA.FileSystem import FileSystem
from config import Config

browser_lib = Selenium()
excel_lib = Files()
pdf_lib = PDF()
table_lib = Tables()
file_lib = FileSystem()


def launchURL(url):
    browser_lib.open_available_browser(url, maximized=True)


def clickLink(locator):
    browser_lib.click_link(locator)


def storeDataInExcel(dict_data, path, worksheet):
    excel_lib.open_workbook(path)
    try:
        excel_lib.append_rows_to_worksheet(dict_data,name= worksheet)
        excel_lib.save_workbook()
    finally:
        excel_lib.close_workbook()


def getAgenciesData():
    agencies_name_keys = []
    agencies_amount_values = []
    agency_name_locator_base = "XPath:(//a//span[contains(@class, 'h4 w200')])"
    agency_amount_locator_base = "XPath:(//a//span[contains(@class, 'h1 w900')])"
    data_count = browser_lib.get_element_count(agency_name_locator_base)
    for i in range(1,data_count+1):
        agency_name_locator = agency_name_locator_base + "[" + str(i) + "]"
        agency_amount_locator = agency_amount_locator_base + "[" + str(i) + "]"
        agencies_name_keys[i] = browser_lib.get_text(agency_name_locator)
        agencies_amount_values[i] = browser_lib.get_text(agency_amount_locator)
        
    agencies_data_dict = {
        "Agency" : agencies_name_keys,
        "Amount" : agencies_amount_values,
    }

    # agencies_data_dict = {agencies_name_keys[i] : agencies_amount_values[i] for i in range(len(agencies_amounts_values))}
    storeDataInExcel(agencies_data_dict, "/output/agencies_data.xlsx", "sheet1")
    # print(agencies_data_dict)


def getTableColumns(locator):
    col_count = browser_lib.get_element_count(locator)
    columns_list = []
    for col in range(1, col_count + 1):
        col_locator = locator + "[" + str(col) + "]"
        columns_list[col] = browser_lib.get_text(col_locator)
    return columns_list


def getTableData(baselocator):
    columns_list = getTableColumns("XPath://*[@id='investments-table-object_wrapper']/div[3]/div[1]/div/table/thead/tr[2]/th")
    col_count = len(columns_list)
    row_count = browser_lib.get_element_count(baselocator)
    data_list = []
    row_data = []
    for row in range(1, row_count+1):
        for col in range(1, col_count + 1):
            cell_locator = "XPath://*[@id='investments-table-object']/tbody/tr[" + str(row) + "]/td[" + str(col) + "]"
            row_data[col] = browser_lib.get_text(cell_locator)
        data_list.append(row_data)
    storeDataInExcel(data_list, "/output/agencies_data.xlsx", "sheet2")


def downloadBusinessCasePDF(baselocator):
    row_count = browser_lib.get_element_count(baselocator)
    for row in range(1, row_count + 1):
        uii_locator = "XPath://*[@id='investments-table-object']/tbody/tr[" + str(row) + "]/td[1]"
        browser_lib.click_link(uii_locator)
        uii = browser_lib.get_text(uii_locator)
        time.sleep(60)
        try:
            browser_lib.click_link("XPath://div//a[text() = 'Download Business Case PDF']")
            time.sleep(60)
            if uii is not None:
                compare_PDFdata_with_webTable(uii,row)
            else:
                break
        finally:
            return "Failed to download PDF"


def compare_PDFdata_with_webTable(uii,row):
    temp_path = file_lib.find_files("*/" + uii)
    file_lib.move_file(temp_path[0],"/output",overwrite=True)
    path = "/output/"+uii+".pdf"
    if file_lib.does_file_exist(path):
        pdf_lib.open_pdf(path)
        text = pdf_lib.get_text_from_pdf(path, [1], trim=False)
        print(text)
        title_excel_value = excel_lib.get_cell_value(row-1,0,"sheet2")
        uii_excel_value = excel_lib.get_cell_value(row-1, 2, "sheet2")
        title_pdf_value = pdf_lib.find_text("text:"+title_excel_value,only_closest=True,trim=False)
        uii_pdf_value = pdf_lib.find_text("text:"+uii,only_closest=True,trim=False)
        pdf_lib.close_all_pdfs()
        if uii_pdf_value[0] == uii_excel_value and title_pdf_value[0] == title_excel_value:
            return True
        else:
            return False


def main():
    try:
        launchURL("https://itdashboard.gov/")
        clickLink("XPath://a[@href='#home-dive-in']")            # Click DIVE-IN
        getAgenciesData()                                        # Scrap Agencies data into excel
        clickLink("XPath:(//a//span[contains(@class, 'h4 w200') and text()='"+Config.Agency_Name+"'])[1]")  # click on any one agency   #Click on Agenecy stored in Config file
        getTableData("XPath://*[@id='investments-table-object']/tbody/tr")               # store Investment table data into excel                                               # Individual Investment table data
        downloadBusinessCasePDF("XPath://*[@id='investments-table-object']/tbody/tr")             # Downloading PDFs
    finally:
        browser_lib.close_all_browsers()


if __name__ == "__main__":
    main()    

