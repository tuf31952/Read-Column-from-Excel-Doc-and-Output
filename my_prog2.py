#!/usr/bin/env python

# for the sheet with 3 names the executtion time was 7.3276 seconds
# for the sheet with 100 names the executtion time was 7.2927 seconds
# for the sheet with 1000 names the executtion time was 7.4259 seconds
# for the sheet with 100000 names the executtion time was 12.6713 seconds but it seems the longer time was due to the time it took to print all the data to stout as opposed to processing the data
# it looks like actual computation time for data processing is around the same no matter how many rows my for loop is parsing through but the time to execution is coming from opening the file

# I seperate each of the sheets out into their own files and the times are:
# for the sheet with 3 names the executtion time was 0.00577 seconds
# for the sheet with 100 names the executtion time was 0.01364 seconds
# for the sheet with 1000 names the executtion time was 0.08659 seconds
# for the sheet with 100000 names the executtion time was 11.5606 seconds
# with each file on its own we can see that each file takes exponently more time to process when opened

import os
import sys
import argparse
import xlsxwriter
import xlrd 
import operator
import time

def main():

    # creats timer to time the execution time of the program
    start_time = time.time()

    # parser to take arguments given in command line as variables and be able to accept files as input
    parser=argparse.ArgumentParser(
    description='''Help page for myprog: ''')
    parser.add_argument('file', type=str, nargs='+', help='Files to be sorted.')
    parser.add_argument('-help', action='store_true', help='Show the help screen.')
    args=parser.parse_args()

    # output the help screen if user enters in -help
    if args.help:
        parser.print_help()
        exit(0)

    # will hold the entries in the names column for output
    names = []

    # append names column to names list array
    for f in args.file:
        wb = xlrd.open_workbook(f)
        sheet = wb.sheet_by_index(0) 
        for i in range(sheet.nrows): 
            names.append(sheet.cell_value(i, 2))

    # sort names list by last name
    names.sort()
    names = sorted(names, key=lambda x: x.split(" ")[-1])

    # print out the names list
    for index, element in enumerate(names):
        print(names[index])

    # print out execution time
    print("--- %s seconds ---" % (time.time() - start_time))


if __name__ == "__main__": 
    main()