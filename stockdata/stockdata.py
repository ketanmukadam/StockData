import pdb
import csv
from bs4 import BeautifulSoup
from selenium import webdriver
from itertools import zip_longest

from SDpickle import SDPickleMixin
from SDscraper import SDScraperMixin
from SDxls import SDxlsMixin


class WgetFinData(SDPickleMixin, SDScraperMixin, SDxlsMixin):
      """
      The class for collecting the 10 year financial data of a company
      """

      def __init__(self, title):
          self.data = []
          self.url = self.query_google(title) 
          if self.url:
             self.driver = webdriver.PhantomJS()
             self.driver.get(self.url)
             #assert title in self.driver.title

      def query_google(self, searchtext):
          """
          query_google() 
          The function to find the stock URL on moneycontrol
          
          Parameters
          ---------
          title: Name of the company to search 
          
          Returns
          -------
          Return the moneycontrol link of the company 
          """
          if not searchtext : return None
          browser = webdriver.PhantomJS()
          browser.get("https://www.google.co.in/search?q="+"moneycontrol "+searchtext)
          assert searchtext in browser.title
          soup = BeautifulSoup(browser.page_source ,"html.parser")
          links = []
          for item in soup.find_all('h3', attrs={'class' : 'r'}):
              links.append(item.a['href'][7:]) # [7:] strips the /url?q= prefix
          browser.quit()
          return links[0]

      def load_pickle_data(self, filename):
          page_list = [1,2,7,8]
          pload = self.pickle_read(filename)
          if not pload:
             for pg in page_list:
                 self.scrape_page(pg)
             self.pickle_dump(filename)




