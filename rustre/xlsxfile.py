#!/usr/bin/env python3
import os.path

from openpyxl import load_workbook
from openpyxl import Workbook

####################################################################################################
# @brief Manage xlsx files
####################################################################################################


class XlsxFile:
    """Manage xlsx file"""

    ####################################################################################################
    # @brief Open a xlsx file
    #   @param[in] filename:
    ####################################################################################################
    def __init__(self, filename):
        self.m_filename = filename
        if not os.path.exists(self.m_filename):
            raise ValueError("File: {} didn't exist!".format(self.m_filename))

        self.m_wb = load_workbook(self.m_filename)
        self.m_sheet = self.m_wb.worksheets[0]
        self.m_total_row = self.m_sheet.max_row
        self.m_total_col = self.m_sheet.max_column

    ####################################################################################################
    # @brief static method for creating empty xlsx file
    #   @param[out] filename  new xlsx filename
    ####################################################################################################
    @staticmethod
    def create_file(filename):
        wb = Workbook()
        page = wb.active
        wb.save(filename)

    ####################################################################################################
    # @brief Get all column values by row number
    #    @param[in] row_number:  Row number (start at 1)
    ####################################################################################################
    def get_columns(self, row_number=1):
        return [cell.value for cell in self.m_sheet[row_number]]

    ####################################################################################################
    # @brief Return the number of rows in the file
    ####################################################################################################
    def get_row_count(self):
        return self.m_total_row

    ####################################################################################################
    # @brief Append a row to the xlsx file
    #    @param[in] row  list of data to append
    ####################################################################################################
    def append_row(self, row):
        self.m_sheet.append(row)

    def save(self):
        self.m_wb.save(self.m_filename)
