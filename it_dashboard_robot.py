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
        self.list_parser = AgenciesListParser()
        logfile = f'{self.output_folder}/log_file.log'
        logging.basicConfig(level=logging.INFO, filename=logfile)

    def run(self):
        # Get agencies
        self.list_parser.click_dive_in_button()
        self.list_parser.get_information()
        # Get individual investments
        agency_link = self.list_parser.get_agency_link(21)
        self.detail_parser = IndividualInvestmentsParser(agency_link)
        details = self.detail_parser.parse()
        # Download PDF
        # Critical: Please move pdf download into its own class
        dir = f'{os.getcwd()}/{self.output_folder}'
        self.browser.set_download_directory(dir, True)
        links = [detail['link'] for detail in details]
        names = [detail['UII'] for detail in details]

        num=0
        for link in links:
            if link != '':
                self.browser.open_available_browser(link)
                # Click to download
                # Critical: Href="#" can be of any button/link element, this is not very efficient, please use a better query
                self.browser.wait_until_element_is_visible('//*[@href="#"]')
                # Low: maybe use click link which can be a better choice here
                self.browser.click_element('//*[@href="#"]')
                # Wait for completed downloads
                # Critical: this while loop here is very bad. It can potentially be a blocking thread if the file can never be download. You will need a timeout 
                while self.lib.does_file_not_exist(
                        '{}/{}.pdf'.format(dir, names[num])):
                    continue
            num += 1
