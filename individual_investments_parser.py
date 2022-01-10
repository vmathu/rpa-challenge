import datetime
from os import path

from RPA.Browser.Selenium import Selenium

from utils import XlsxSaver


class IndividualInvestmentsParser:

    def __init__(self, agency_link):
        self.browser = Selenium()
        self.browser.open_available_browser(agency_link)

    def _get_investments_table(self):
        self.browser.wait_until_element_is_visible(
            locator='//select[@class="form-control c-select"]/option[4]',
            timeout=datetime.timedelta(seconds=60)
        )
        self.browser.click_element(
            locator='//select[@class="form-control c-select"]/option[4]'
        )
        self.browser.wait_until_element_is_visible(
            locator='//a[@class="paginate_button last disabled"]',
            timeout=datetime.timedelta(seconds=60)
        )

    def _get_table_headers(self):
        headers = self.browser.get_webelements(
            '//*[@class="dataTables_scrollHead"]/div/table/thead/tr[2]/th')
        return [header.text for header in headers]

    def _get_table_rows(self):
        self.browser.wait_until_element_is_enabled(
            locator='//table[@id="investments-table-object"]/tbody/tr'
        )
        return self.browser.get_webelements(
            locator='//table[@id="investments-table-object"]/tbody/tr'
        )

    def _get_row_cells(self, table_row):
        return self.browser.get_webelements(locator=[table_row, 'css:td'])

    def _save_investments_table(self, headers, investments):
        helper = XlsxSaver(
            path='output/Agencies.xlsx',
            worksheet='Individual Investments'
        )
        helper._fill_headers(2, headers)
        helper._fill_rows(investments)
        helper._save_workbook()

    def _get_investments(self):
        self._get_investments_table()
        headers = self._get_table_headers()
        rows = self._get_table_rows()
        investments = []

        self.browser.wait_until_element_is_visible(
            locator='//table[@id="investments-table-object"]/tbody/tr/td/a',
            timeout=datetime.timedelta(seconds=10)
        )
        for row in rows:
            investment = {}
            cells = self._get_row_cells(row)
            num = 0
            for cell in cells:
                if num == 0:
                    count_a = self.browser.get_element_count(
                        locator=[cell, 'css:a'])
                    if count_a > 0:
                        investment["link"] = self.browser.get_element_attribute(
                            locator=[cell, 'css:a'],
                            attribute='href'
                        )
                    else:
                        investment["link"] = ''
                investment[headers[num]] = cell.text
                num += 1
            investments.append(investment)
        self._save_investments_table(headers, investments)
        return investments
