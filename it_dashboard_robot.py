import logging
import os

from RPA.Browser.Selenium import Selenium
from RPA.FileSystem import FileSystem

from agencies_list_parser import AgenciesListParser
from individual_investments_parser import IndividualInvestmentsParser


class ITDashboardRobot:

    def __init__(self):
        self.browser = Selenium()
        self.lib = FileSystem()
        self.output_folder = 'output'
        self.list_parser = AgenciesListParser("https://itdashboard.gov")
        logfile = f'{self.output_folder}/log_file.log'
        logging.basicConfig(level=logging.INFO, filename=logfile)

    def run(self):
        # Get agencies
        self.list_parser._click_dive_in_button()
        self.list_parser._get_information()
        # Get individual investments
        agency_link = self.list_parser._get_agency_link(21)
        self.detail_parser = IndividualInvestmentsParser(agency_link)
        details = self.detail_parser._get_investments()
        # Download PDF
        dir = f'{os.getcwd()}/{self.output_folder}'
        self.browser.set_download_directory(dir, True)
        links = [detail['link'] for detail in details]
        names = [detail['UII'] for detail in details]

        num = 0
        for link in links:
            if link != '':
                self.browser.open_available_browser(link)
                # Click to download
                self.browser.wait_until_element_is_visible(
                    '//*[@id="business-case-pdf"]/a')
                self.browser.click_link(
                    '//*[@id="business-case-pdf"]/a')
                # Wait for completed downloads
                self.lib.wait_until_created(
                    f'{dir}/{names[num]}.pdf',
                    timeout=60.0*5
                )
            num += 1
