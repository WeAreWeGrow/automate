from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem
from RPA.PDF import PDF
import time, copy

file_system = FileSystem()
selenium = Selenium()
files = Files()
pdf = PDF()


def act_on_element(path, action):
    success = False
    while not success:
        try:
            if action == "click_element":
                selenium.click_element(path)
                success = True
            elif action == "find_elements":
                results = selenium.find_elements(path)
                if results:
                    return results
                else:
                    raise Exception
        except:
            time.sleep(2)

def get_selected_agency():
    selected_agency = file_system.read_file("selected_agency.txt", encoding = 'utf-8')
    return selected_agency

def get_agencies_information(selected_agency):
    preferences_dict = {"profile.default_content_settings.popups": 0, "directory_upgrade": True}
    #Local
    # preferences_dict["download.default_directory"] = r"C:\Users\carlo\Documents\Programming-Projects\Python Scripts\RPAITDashboard\\"
    #Robocloud
    # route = file_system.absolute_path("output") + "\\"
    # preferences_dict["download.default_directory"] = r'%s' % route
    #Robocloud Workforce Process
    preferences_dict["download.default_directory"] = "output/"
    selenium.open_available_browser("https://itdashboard.gov/", maximized = True, preferences = preferences_dict)
    act_on_element("xpath://a[@class='btn btn-default btn-lg-2x trend_sans_oneregular']", "click_element")
    general_path = "xpath://div[@id='agency-tiles-container']//div[@class='row top-margin-10 top-gutter-20 dash-bottom']//div[@class='col-sm-12']"
    agencies_names = act_on_element(general_path + "//span[@class='h4 w200']", "find_elements")
    agencies_values = act_on_element(general_path + "//span[@class=' h1 w900']", "find_elements")
    agencies_buttons = act_on_element(general_path + "//a[@class='btn btn-default btn-sm']", "find_elements")
    selected_agency_button = None
    agencies_data_list = []
    for agency_name, agency_value, agency_button in zip(agencies_names, agencies_values, agencies_buttons):
        agency_data_dict = {"Agency": agency_name.text, "Amount": agency_value.text}
        agencies_data_list.append(agency_data_dict)
        if agency_name.text == selected_agency:
            selected_agency_button = agency_button
    return agencies_data_list, selected_agency_button

def get_selected_agency_details(selected_agency_button):
    selenium.click_element(selected_agency_button)
    act_on_element("xpath://select[@class='form-control c-select']//option[@value='-1']", "click_element")
    table_titles = act_on_element("xpath://div[@class='dataTables_scrollHead']//tr[@role='row']//th[@aria-controls='investments-table-object']", "find_elements")
    table_titles_length = len(table_titles)
    selected_agency_data_list = []
    table_datas = act_on_element("xpath://div[@class='dataTables_scrollBody']//tr[@role='row']//td", "find_elements")
    table_links = act_on_element("xpath://div[@class='dataTables_scrollBody']//tr[@role='row']//td//a", "find_elements")
    for index, table_link in enumerate(table_links):
        table_links[index] = table_link.get_attribute('href')
    table_datas = [table_datas[i:i + table_titles_length] for i in range(0, len(table_datas), table_titles_length)]
    for table_data in table_datas:
        selected_agency_data_dict = {}
        for data, table_title in zip(table_data, table_titles):
            selected_agency_data_dict[table_title.text] = data.text
        selected_agency_data_list.append(selected_agency_data_dict)
    return table_links, selected_agency_data_list

def write_excel(selected_agency, agencies_data_list, selected_agency_data_list):
    #Local
    # files.create_workbook(path = 'IT_Dashboard.xlsx', fmt = 'xlsx')
    #Robocloud
    files.create_workbook(path = 'output/IT_Dashboard.xlsx', fmt = 'xlsx')
    files.create_worksheet(name = "Agencies", content = None, exist_ok = False, header = False)
    files.create_worksheet(name = selected_agency, content = None, exist_ok = False, header = False)
    files.append_rows_to_worksheet(agencies_data_list, name = "Agencies", header = True, start = None)
    files.append_rows_to_worksheet(selected_agency_data_list, name = selected_agency, header = True, start = None)
    files.remove_worksheet(name = "Sheet")
    files.save_workbook(path = None)
    files.close_workbook()

def download_pdfs(table_links):
    downloaded_pdfs_amount = 0
    for table_link in table_links:
        selenium.go_to(table_link)
        act_on_element("xpath://div[@id='business-case-pdf']//a", "click_element")
        actual_downloaded_pdfs_amount = copy.copy(downloaded_pdfs_amount)
        while not downloaded_pdfs_amount == actual_downloaded_pdfs_amount + 1:
            #Local
            # downloaded_pdfs_amount = len(file_system.find_files("*.pdf"))
            #Robocloud
            downloaded_pdfs_amount = len(file_system.find_files("output/*.pdf"))
            time.sleep(2)

def compare_values(selected_agency_data_list):
    #Local
    # matches = file_system.find_files("*.pdf")
    #Robocloud
    matches = file_system.find_files("output/*.pdf")
    pdf_data_dict = {}
    for match in matches:
        text = pdf.get_text_from_pdf(match)
        text = text[1].split('(All Capital Assets)')[1].split('Section B')[0]
        investment_name = text.split("Name of this Investment: ")[1].split('2. Unique Investment Identifier (UII)')[0]
        uii = text.split('Unique Investment Identifier (UII): ')[1]
        pdf_data_dict[uii] = investment_name
    comparison_data = ""
    for selected_agency_data in selected_agency_data_list:
        if selected_agency_data["UII"] in pdf_data_dict:
            comparison_data = comparison_data + "UII (Webpage): " + selected_agency_data["UII"] + "\n"
            comparison_data = comparison_data + "Unique Investment Identifier (UII) (PDF): " + selected_agency_data["UII"] + "\n"
            comparison_data = comparison_data + "Investment Title (Webpage): " + selected_agency_data["Investment Title"] + "\n"
            comparison_data = comparison_data + "Name of this Investment (PDF): " + pdf_data_dict[selected_agency_data["UII"]] + "\n"
            comparison_data = comparison_data + "--------------------------------------" + "\n"
    #Local
    # file_system.create_file('Comparison.txt', content = comparison_data, encoding = 'utf-8', overwrite = True)
    #Robocloud
    file_system.create_file('output/Comparison.txt', content = comparison_data, encoding = 'utf-8', overwrite = True)


def main():
    try:
        selected_agency = get_selected_agency()
        agencies_data_list, selected_agency_button = get_agencies_information(selected_agency)
        table_links, selected_agency_data_list = get_selected_agency_details(selected_agency_button)
        write_excel(selected_agency, agencies_data_list, selected_agency_data_list)
        download_pdfs(table_links)
        compare_values(selected_agency_data_list)
    finally:
        selenium.close_all_browsers()

if __name__ == "__main__":
    main()
