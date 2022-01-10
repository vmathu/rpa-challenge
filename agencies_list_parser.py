from RPA.Browser.Selenium import Selenium

from utils import XlsxSaver


class AgenciesListParser:

    def __init__(self, url):
        self.browser = Selenium()
        self.browser.open_available_browser(url)

    def _click_dive_in_button(self):
        self.browser.wait_until_element_is_visible(
            '//*[@href="#home-dive-in"]'
        )
        self.browser.click_element('//*[@href="#home-dive-in"]')

    def _get_agencies(self):
        self.browser.wait_until_element_is_visible(
            '//*[@id="agency-tiles-container"]'
        )
        return self.browser.get_webelements(
            '//*[@id="agency-tiles-container"]/div/div/div/div')

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

    def _get_information(self):
        agencies = self._get_agencies()
        helper = XlsxSaver(
            worksheet='Agencies',
            path='output/Agencies.xlsx'
        )
        row = 1
        for agency in agencies:
            spending = self._get_spending(agency).text
            helper._fill_row(row, "A", spending)
            row += 1
        helper._save_workbook()

    def _get_agency_link(self, id):
        agency = self._get_agencies()[id]
        link = self.browser.get_webelement(
            locator=[agency, 'css:div>div>div>a']
        )
        return link.get_attribute('href')
