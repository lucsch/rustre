#!/usr/bin/env python3
import logging
import configparser
from rustre.xlsxfile import XlsxFile

####################################################################################################
# @brief Compare two xlsx files
####################################################################################################


class XlsxCompare:
    """Compare two xlsx"""

    ####################################################################################################
    # @brief Init the XlsxCompare object
    #    @param[in] config_file  config file (ini)
    #    @param[in] file_source  source file (xlsx)
    #    @param[in] file_target  target file (xlsx)
    ####################################################################################################
    def __init__(self, config_file, file_source, file_target):
        self.m_config_file = config_file
        self.m_file_source = file_source
        self.m_file_target = file_target

    def do_compare(self, log_file):
        # open the config file
        pass




