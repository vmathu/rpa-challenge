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
        self.user_director = os.path.expanduser('~')

    def run(self):
        #Get agencies
        self.list_parser.click_dive_in_button()
        self.list_parser.get_information()
        #Get individual investments
        agency_link = self.list_parser.get_agency_link()
        self.detail_parser = IndividualInvestmentsParser(agency_link)
        details = self.detail_parser.parse()
        #Download PDF
        links = [detail['link'] for detail in details]
        names = [detail['UII'] for detail in details]

        num=0
        for link in links:
            if link != '':
                self.browser.open_available_browser(link)
                #Click to download
                self.browser.wait_until_element_is_visible('//*[@href="#"]')
                self.browser.click_element('//*[@href="#"]')
                #Wait for completed downloads
                while self.lib.does_file_not_exist('{}/Downloads/{}.pdf'.format(self.user_director, names[num])):
                    continue     
                else: 
                    #Move file from default download folder to ./output
                    self.lib.move_file('{}/Downloads/{}.pdf'.format(self.user_director, names[num]), '{}/output/{}.pdf'.format(os.getcwd(), names[num]), True)
            num += 1