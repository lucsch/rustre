#!/usr/bin/env python3
import os
from rustre.xlsxfile import XlsxFile

class XlsxDuplicate:
    """
    Check a xlsx file for duplicates.
    :param src_filename:
    :param log_filename:
    :param cols:
    """
    def __init__(self, src_filename :str, log_filename :str, cols :[], sheet_index=0, header_index=1):
        self.m_src_filename = src_filename
        self.m_log_filename = log_filename
        self.m_cols = cols
        self.m_sheet_index = sheet_index
        self.m_header_index = header_index

    def _is_valid(self) -> bool:
        if not os.path.exists(self.m_src_filename):
            return False
        if self.m_log_filename == "":
            return False
        if len(self.m_cols) == 0:
            return False
        return True

    def check_duplicate(self) -> bool:
        if not self._is_valid():
            return False

        # open log file
        XlsxFile.create_file(self.m_log_filename)
        xlsx_log = XlsxFile(self.m_log_filename)
        xlsx_log.append_row(["ID", "COUNT"])

        # open source file
        xlsx_src = XlsxFile(self.m_src_filename, self.m_sheet_index, load_in_memory=True)

        xlsx_log.save()
        return True

