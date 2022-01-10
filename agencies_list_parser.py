from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
# Critical: Don't import Excel everywhere. It defeat the purpose of using class and encapsulation. You should create an Excel writer (with methods for 
# writing different files) and then import that class in here.  

class AgenciesListParser:
    URL = "https://itdashboard.gov"

    # Medium: Please move this into where is it use, it is only used once. It is bad to have global variable like this. And the name of this variable 
    # is not descriptive. Don't use generic words for variable name.
    array = []

    def __init__(self):
        self.browser = Selenium()
        self.browser.open_available_browser(self.URL)
        self.excel = Files()

    def click_dive_in_button(self):
        self.browser.wait_until_element_is_visible(
            '//*[@href="#home-dive-in"]'
        )
        self.browser.click_element('//*[@href="#home-dive-in"]')

    def get_agencies(self):
        self.browser.wait_until_element_is_visible(
            '//*[@id="agency-tiles-container"]'
        )
        container = self.browser.find_element(
            '//*[@id="agency-tiles-container"]'
        )
        return container.find_elements_by_class_name('noUnderline')
    # Critical: Never, ever use styling classname to find an element. Styling can change. However structure will not. Please use structure like the example from 
    # https://github.com/kvazarich/IT_Dashboard_RPA_Challenge

    def _get_name(self, agency):
        self.browser.wait_until_element_is_visible(
            locator=[agency, 'css:span:nth-of-type(1)']
        )
        return self.browser.get_webelement(
            locator=[agency, 'css:span:nth-of-type(1)']
        )

    def _get_spending(self, agency):
        self.browser.wait_until_element_is_visible(
            locator=[agency, 'css:span:nth-of-type(2)']
        )
        return self.browser.get_webelement(
            locator=[agency, 'css:span:nth-of-type(2)']
        )
# Medium: Your naming convention is really inconsistent, please unify them by whether or not to put a _ at the start of methods names

    def get_information(self):
        agencies = self.get_agencies()
        # Critical: Move this entire writing method into an ExcelWriter class
        self.excel.create_workbook("output/Agencies.xlsx")
        self.excel.rename_worksheet("Sheet", "Agencies")
        row=1
        for agency in agencies:
            name = self._get_name(agency).text
            spending = self._get_spending(agency).text
            self.array.append(
                {"name": name, "spending": spending}
            )
            self.excel.set_cell_value(row, "A", spending)
            row += 1
        self.excel.save_workbook()
    
    def get_agency_link(self, id):
        agency = self.get_agencies()[id]
    # Critical: Never, ever use styling classname to find an element. Styling can change. However structure will not. Please use structure like the example from 
    # https://github.com/kvazarich/IT_Dashboard_RPA_Challenge
        return agency.find_element_by_class_name('btn-sm').get_attribute('href')
