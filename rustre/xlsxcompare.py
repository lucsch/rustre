#!/usr/bin/env python3
import logging
import os.path
from rustre.xlsxfile import XlsxFile

####################################################################################################
# @brief Compare two xlsx files
# @details the comparison is made on a column basis
####################################################################################################


class XlsxCompare:
    """Compare two xlsx headers"""

    ####################################################################################################
    # @brief Init the XlsxCompare object
    #    @param[in] file1  First file (xlsx only)
    #    @param[in] file2  Second file (xlsx only)
    ####################################################################################################
    def __init__(self, file1, file2, sheet_index1=0, sheet_index2=0):
        self.m_file1 = file1
        self.m_file2 = file2
        self.m_xlsx1 = XlsxFile(self.m_file1, sheet_index1)
        self.m_xlsx2 = XlsxFile(self.m_file2, sheet_index2)

    ####################################################################################################
    # @brief Compare headers in two files
    #    @param[in] row_number_file1=1  Row number for first file (start at 1)
    #    @param[in] row_number_file2=1  Row number for second file (start at 1)
    ####################################################################################################
    def compare_headers(self, row_number_file1=1, row_number_file2=1):
        my_cols1 = self.m_xlsx1.get_columns(row_number_file1)
        my_cols2 = self.m_xlsx2.get_columns(row_number_file2)
        if len(my_cols1) == 0:
            logging.error("no columns in file: {}".format(self.m_file1))
        if len(my_cols2) == 0:
            logging.error("no columns in file: {}".format(self.m_file2))
        if my_cols1 == my_cols2:
            return True
        logging.error("Columns {}: {}".format(self.m_file1, my_cols1))
        logging.error("Columns {}: {}".format(self.m_file2, my_cols2))
        return False


