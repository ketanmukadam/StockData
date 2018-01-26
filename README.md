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

## License

See the [LICENSE](LICENSE.md) file for license rights and limitations.
