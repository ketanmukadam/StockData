
import re
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException 

bs_keyref = ['EQUITIES AND LIABILITIES',"SHAREHOLDER'S FUNDS",'Total Share Capital',
    'Total Reserves and Surplus', 'NON-CURRENT LIABILITIES',
    'Total Non-Current Liabilities', 'Total Current Liabilities',
    'Total Capital And Liabilities', 'NON-CURRENT ASSETS',
    'Fixed Assets', 'Total Non-Current Assets',
    'Total Current Assets', 'Total Assets',
    'CONTINGENT LIABILITIES, COMMITMENTS', 'CIF VALUE OF IMPORTS',
    'EXPENDITURE IN FOREIGN EXCHANGE', 'REMITTANCES IN FOREIGN CURRENCIES FOR DIVIDENDS',
    'EARNINGS IN FOREIGN EXCHANGE', 'BONUS DETAILS',
    'NON-CURRENT INVESTMENTS', 'CURRENT INVESTMENTS','Source :  Dion Global Solutions Limited']

pl_keyref = ['Revenue From Operations [Gross]','Revenue From Operations [Net]',
             'Total Operating Revenues','EXPENSES','Total Expenses',
             'Profit/Loss Before Tax','Total Tax Expenses','Profit/Loss For The Period',
             'EARNINGS PER SHARE','VALUE OF IMPORTED AND INDIGENIOUS RAW MATERIALS',
             'DIVIDEND AND DIVIDEND PERCENTAGE','Equity Dividend Rate (%)','Source :  Dion Global Solutions Limited']

cf_keyref = ['Net Profit/Loss Before Extraordinary Items And Tax', 
             'Net Inc/Dec In Cash And Cash Equivalents',
             'Source :  Dion Global Solutions Limited']

ratio_keyref = ['Per Share Ratios', 'Profitability Ratios', 
                'Liquidity Ratios', 'Valuation Ratios','Source :  Dion Global Solutions Limited']


class SDScraperMixin():
      def get_list_idx(self, key, data):
          for idx, element in enumerate(data):
              if key in element: return idx
      def sort_list(self, data1,data2):
          data = data1[1:] + data2[1:]
          data = list(filter(None,data))
          data.sort(key=lambda x: x[0])
          data = data1[:1] + data2[:1] + data
          return data
              
      def merge_sublists(self, data1, data2):
          data = self.sort_list(data1,data2) 
          data_iter = enumerate(data)
          for i,row in data_iter:
               if row != data[-1] and data[i][0] == data[i+1][0]:
                  if data[i] == data[i+1]:
                     self.data.append(data[i])
                  else:
                     self.data.append(data[i] + data[i+1][1:])
                  next(data_iter)
               else:
                  if i != len(data)-1 and data[i] == data[i+1]:
                     continue
                  if len(row) == 1 : 
                     self.data.append(row)
                     continue
                  if row in data1 :
                     self.data.append(row + ['0.00'] * 5)
                  else :
                     self.data.append(row[:1] + ['0.00'] * 5 + row[1:])

      def cleanze_data(self, data):
          first = True
          data = list(filter(None,data))
          if len(data) < 5 : return None
          for row in data[:]:
              if any(re.search('12 mths',ele) is not None for ele in row):
                    data.remove(row)
                    continue
              if any(re.search('Mar \d{2}',ele) is not None for ele in row):
                 if first:
                    row.insert(0,'YEAR')
                    first = False
                 else:
                    data.remove(row)
          return data

      def change_data_type(self):
          for row in self.data:
              for idx, ele in enumerate(row):
                  #if not idx: continue
                  try:
                    if type(row[idx]) != float:
                       row[idx] = float(row[idx].replace(',','')) 
                  except ValueError: 
                    pass

      def merge_lists(self, data1, data2, num):
          if num == 1 : keyref = bs_keyref
          if num == 2 : keyref = pl_keyref
          if num == 7 : keyref = cf_keyref
          if num == 8 : keyref = ratio_keyref 

          data1 = self.cleanze_data(data1)
          data2 = self.cleanze_data(data2)

          if not data1 and data2: return 
          if not data1 : data1 = data2
          if not data2 : data2 = data1

          previdx1 = 0
          curidx1 = 0
          previdx2 = 0
          curidx2 = 0
          for ele in keyref:
              curridx1 = self.get_list_idx(ele, data1)
              curridx2 = self.get_list_idx(ele, data2)
              if not curridx1 or not curridx2 : 
                 print('%s not found' % ele)  
              self.merge_sublists(data1[previdx1:curridx1],data2[previdx2:curridx2])
              previdx1 = curridx1
              previdx2 = curridx2
      def pull_data(self, page_source):
          data = []
          soup = BeautifulSoup(page_source,"lxml")
          divTags = soup.find_all('div', {'class':'boxBg'})

          for tag in divTags:
               tables = tag.find_all('table',attrs={'class':'table4'})
               for table in tables:
                   table_bodys = table.find_all('tbody')
                   for tablebody in table_bodys:
                       rows = tablebody.find_all('tr')
                       for row in rows:
                           cols = row.find_all('td')
                           cols = [ele.text.strip() for ele in cols]
                           data.append([ele for ele in cols if ele]) # Get rid of empty values
          return data
      def scrape_page(self,num):
          navigXpath = ['''//*[@id="slider"]/dt[7]/a''',
                '''//*[@id="slider"]/dd[3]/ul/li['''+str(num)+''']/a''',
		'''//*[@id="mc_mainWrapper"]/div[2]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a/b''']
          try:
             self.driver.find_element_by_xpath(navigXpath[0]).click()
             self.driver.find_element_by_xpath(navigXpath[1]).click()
             yr1 = self.pull_data(self.driver.page_source)
             self.driver.find_element_by_xpath(navigXpath[2]).click()
             yr2 = self.pull_data(self.driver.page_source)
             self.merge_lists(yr1, yr2, num)
             self.change_data_type()
          except NoSuchElementException:
             print ("Webpage Not Accessible, Try again after some time")
