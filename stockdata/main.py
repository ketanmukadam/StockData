#!/usr/bin/python3 
import stockdata
import sys
import os
from optparse import OptionParser

def main(options):
    """
    main() function to start the program 
    
    Parameters
    ---------
    options : Command line arguments parsed by OptionParser 
    
    Returns
    -------
    None 
    """
    xlsinfile = "out/template.xlsx"
    worksheetname = "InputData"
    checkdir(xlsinfile)
    u = stockdata.WgetFinData(options.searchtext)
    url = u.url
    print("URL:-> "+url)
    xlsoutfile = 'out/' + url.split(sep='/')[6] + '.xlsx'
    xlsfullfile = 'out/' + url.split(sep='/')[6] + 'full.xlsx'
    u.load_pickle_data('out/' + url.split(sep='/')[6]+'pickle.p')
    u.copy_fulldata(xlsfullfile)
    u.copy_mysheet(xlsinfile, xlsoutfile, worksheetname)
    print("Copied Data to %s" % xlsoutfile)

def checkdir(filename):
    if not os.path.exists(os.path.dirname(filename)):
       try:
          os.makedirs(os.path.dirname(filename))
       except OSError as exc: # Guard against race condition
          if exc.errno != errno.EEXIST:
             raise

def argparser():
    """
    Option parsing of command line

    It will add the required arguments to OptionParser module
    Collects and parse the arguments
    
    Parameters
    ---------
    None 
    
    Returns
    -------
    opts: Parsed arguments (or their defaults) returned in opts
    """
    parser = OptionParser(usage="usage: %prog [options]")
    parser.add_option(
        "-s","--srchtxt", dest="searchtext", help="text for search", default="Mayur Uniquoter")
    (opts, args) = parser.parse_args(sys.argv)
    return opts

if __name__ == '__main__':
     # Parse the arguments and pass to main
     options = argparser()
     sys.exit(main(options))


