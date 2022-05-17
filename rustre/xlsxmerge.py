#!/usr/bin/env python3
import logging
import os.path
from rustre.xlsxfile import XlsxFile
from rustre.xlsxcompare import XlsxCompare


####################################################################################################
# @brief Merge multiple xlsx files
# @details Check is done to ensure files are similar
####################################################################################################
class XlsxMerge:
    """Merge multiple xlsx files"""

    ####################################################################################################
    # @brief Init XlsxMerge object
    #   @param[in] source_files  list of xlsx filenames
    #   @param[in] sheet_index  index of the sheet
    #   @param[in] header_index  index of the header
    ####################################################################################################
    def __init__(self, source_files, sheet_index=0, header_index=1):
        self.m_source_files = source_files
        self.m_sheet_index = sheet_index
        self.m_header_index = header_index
        if source_files is None or len(source_files) == 0:
            raise ValueError("No source files")

    ####################################################################################################
    # @brief Merge the files into the result filename
    #    @param[out] result_file  xlsx output filename
    ####################################################################################################
    def merge(self, result_file):
        # compare headers
        for index in range(1, len(self.m_source_files)):
            xcomp = XlsxCompare(self.m_source_files[0], self.m_source_files[index], self.m_sheet_index, self.m_sheet_index)
            if not xcomp.compare_headers(self.m_header_index, self.m_header_index):
                logging.error("Xlsx files are not similar!")
                return False

        # merge first file in the new filename
        XlsxFile.create_file(result_file)
        if not os.path.exists(result_file):
            logging.error("Unable to create: {}".format(result_file))
            return False
        xlsx_result = XlsxFile(result_file)
        xlsx_source1 = XlsxFile(self.m_source_files[0], self.m_sheet_index)
        m_row_tot = xlsx_source1.get_row_count()
        for r in range(self.m_header_index, m_row_tot+1):
            my_row = xlsx_source1.get_columns(r)
            xlsx_result.append_row(my_row)

        # merge the other files
        for index in range(1, len(self.m_source_files)):
            xlsx_src = XlsxFile(self.m_source_files[index], self.m_sheet_index)
            for r in range(self.m_header_index+1, xlsx_src.get_row_count()+1):
                my_row = xlsx_src.get_columns(r)
                xlsx_result.append_row(my_row)
        xlsx_result.save()
        return True





