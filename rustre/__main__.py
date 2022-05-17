#!/usr/bin/env python3
import os.path
import sys
from rustre.xlsxmerge import XlsxMerge
from rustre.xlsxfile import XlsxFile

if __name__ == '__main__':
    if len(sys.argv) <= 1:
        print("Rustre usage:python -m rustre file.xlsx file2.xlsx ...")
        exit()

    filenames = sys.argv[1:]
    # xlsx = XlsxFile(filenames[0], 1)
    # print(xlsx.get_columns(1))
    #
    # xlsx = XlsxFile(filenames[1], 1)
    # print(xlsx.get_columns(1))
    #
    # xlsx = XlsxFile(filenames[2], 1)
    # print(xlsx.get_columns(1))
    # exit()

    xmrg = XlsxMerge(filenames, sheet_index=0, header_index=7)
    res_file = os.path.expanduser("~/result.xlsx")
    xmrg.merge(res_file)
    print("export done in: '{}'".format(res_file))



