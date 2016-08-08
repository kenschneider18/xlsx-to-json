import sys
import getopt
import openpyxl
import simplejson

from collections import OrderedDict


def main(argv):
    inputfile = ''
    outputfile = ''
    sheetnumber = 0

    try:
        opts, args = getopt.getopt(argv, "hsi:o:", ["help", "ifile", "ofile", "sheet"])
    except getopt.GetoptError:
        print 'app.py -i <inputfile> -o <outputfile> [-s <sheetnumber>]'
        sys.exit(2)
    for opt, arg in opts:
        if opt in ('-h', '--help'):
            print 'app.py -i <inputfile> -o <outputfile> [-s <sheetnumber>]'
            sys.exit()
        elif opt in ('-i', '--ifile'):
            inputfile = arg
        elif opt in ('-o', '--ofile'):
            outputfile = arg
        elif opt in ('-s', '--sheet'):
            try:
                sheetnumber = int(arg)
            except ValueError:
                print 'app.py -i <inputfile> -o <outputfile> [-s <sheetnumber>]'
                sys.exit(2)


    excel = openpyxl.load_workbook(inputfile)
    sheet = excel.sheet_by_index(sheetnumber)

    json_list = []
