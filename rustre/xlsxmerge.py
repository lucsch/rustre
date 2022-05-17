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
    ####################################################################################################
    def __init__(self, source_files):
        self.m_source_files = source_files

    ####################################################################################################
    # @brief Merge the files into the result filename
    #    @param[out] result_file  xlsx output filename
    ####################################################################################################
    def merge(self, result_file):
        # compare headers
        for index in range(1, len(self.m_source_files)):
            xcomp = XlsxCompare(self.m_source_files[0], self.m_source_files[index])
            if not xcomp.compare_headers(1, 1):
                logging.error("Xlsx files are not similar!")
                return False



        return True





