import datetime
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files


class IndividualInvestmentsParser:

    def __init__(self, agency_link):
        self.browser = Selenium()
        self.browser.open_available_browser(agency_link)
        self.excel = Files()

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
        self.browser.get_webelement(
            locator='//*[@id="investments-table-widget"]'
        )

    def _get_table_headers(self):
        self.browser.wait_until_element_is_visible(
            locator='//*[@class="dataTables_scrollHead"]/div/table/thead/tr[2]/th',
            timeout=datetime.timedelta(seconds=60)
        )
        headers = self.browser.get_webelements(
            locator='//*[@class="dataTables_scrollHead"]/div/table/thead/tr[2]/th'
        )
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

    def save_investments_table(self, headers, investments):
        self.excel.open_workbook('output/Agencies.xlsx')
        self.excel.create_worksheet('Individual Investments')
        col = 2
        for header in headers:
            self.excel.set_cell_value(1, col, header)
            col += 1
        self.excel.append_rows_to_worksheet(investments)
        self.excel.save_workbook()

    def parse(self):
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
        self.save_investments_table(headers, investments)
        return investments
