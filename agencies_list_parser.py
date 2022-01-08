from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files


class AgenciesListParser:
    URL = "https://itdashboard.gov"

    array = []

    def __init__(self):
        self.browser = Selenium()
        self.browser.open_available_browser(self.URL)
        self.excel = Files()

    def click_dive_in_button(self):
        self.browser.wait_until_element_is_visible('//*[@href="#home-dive-in"]')
        self.browser.click_element('//*[@href="#home-dive-in"]')

    def get_agencies(self):
        self.browser.wait_until_element_is_visible('//*[@id="agency-tiles-container"]')
        container = self.browser.find_element('//*[@id="agency-tiles-container"]')
        return container.find_elements_by_class_name('noUnderline')

    def _get_name(self, agency):
        self.browser.wait_until_element_is_visible(locator=[agency, 'css:span:nth-of-type(1)'])
        return self.browser.get_webelement(locator=[agency, 'css:span:nth-of-type(1)'])

    def _get_spending(self, agency):
        self.browser.wait_until_element_is_visible(locator=[agency, 'css:span:nth-of-type(2)'])
        return self.browser.get_webelement(locator=[agency, 'css:span:nth-of-type(2)'])

    def get_information(self):
        agencies = self.get_agencies()
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
    
    def get_agency_link(self):
        return self.get_agencies()[18].find_element_by_class_name('btn-sm').get_attribute('href')