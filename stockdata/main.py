#!/usr/bin/python3 
import stockdata
import sys

url = "http://www.moneycontrol.com/india/stockpricequote/plastics/mayuruniquoters/MU"
#url = "http://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/sunpharmaceuticalindustries/SPI"

xlsinfile = "Stock_Analysis_Final.xlsx"
xlsoutfile = url.split(sep='/')[-2] + '.xlsx'
xlsfullfile = url.split(sep='/')[-2] + 'full.xlsx'
worksheetname = "InputData"

def main:
    u = WgetFinData(url,"Mayur")
    u.get_data(url.split(sep='/')[-2]+'pickle.p')
    #u.print_table()
    u.copy_fulldata(xlsfullfile)
    u.copy_mysheet(xlsinfile, xlsoutfile, worksheetname)
    print("Copied Data to %s" % xlsoutfile)


if __name__ == '__main__':
     sys.exit(main())


