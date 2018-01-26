# Stock Data Script

This repository contains two projects. 
* `StockData` is a python application to download Indian Company's `Financial Data` from [Moneycontrol website](http://www.moneycontrol.com). It downloads the 10 years balance sheet, profit & loss statement and financial ratios as published on moneycontrol. The application will format and arrange the data in an excel sheet for further use. 

          Usage: main.py [options]
          Options:
            -h, --help                            show this help message and exit
            -s SEARCHTEXT, --srchtxt=SEARCHTEXT   text for search
* `GetFinancials` is another python script for downloading Indian Company's `Financial Data` from [Moneycontrol website](http://www.moneycontrol.com). It uses *pandas library* to manipulate data and save in the excel template format. 

          Usage: getdata.py [options]
          Options:
            -h, --help                                   show this help message and exit
            -u URL, --url=URL                            Moneycontrol URL for stock
            -t TMPLTXLSX, --template=TMPLTXLSX           Template xlsx file
            -o OUTTMPLTXLSX, --outtemplate=OUTTMPLTXLSX  Out Template xlsx file
            -x INPUTXLSX, --inputxls=INPUTXLSX           Input xlsx file
            -c, --consolidateflag                        Get consolidated results

Example output of excel sheet
-----------------------------
* Balance Sheet
<img width="1006" alt="screenshot at jan 26 14-11-32" src="https://github.com/ketanmukadam/StockData/blob/master/store/BalanceSheet.png">

* Profit and Loss 
<img width="806" alt="screenshot at jan 26 14-11-32" src="https://github.com/ketanmukadam/StockData/blob/master/store/ProfitLoss.png">

* Financial Ratios 
<img width="806" alt="screenshot at jan 26 14-11-32" src="https://github.com/ketanmukadam/StockData/blob/master/store/Ratios.png">


## License

See the [LICENSE](LICENSE.md) file for license rights and limitations.
