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

        # open source file
        xlsx_src = XlsxFile(self.m_src_filename, self.m_sheet_index, load_in_memory=True)

        # open log file
        XlsxFile.create_file(self.m_log_filename)
        xlsx_log = XlsxFile(self.m_log_filename)
        header = self._get_col_id(xlsx_src.get_columns(self.m_header_index))
        header.append("COUNT")
        xlsx_log.append_row(header)

        # iterate
        skip_index = []
        for row_index in range(self.m_header_index+1, xlsx_src.get_row_count() + 1):
            if row_index in skip_index:
                continue
            row = xlsx_src.get_columns(row_index)
            id_val = self._get_col_id(row)
            count = 0
            for row_check_index in range(row_index, xlsx_src.get_row_count() + 1):
                row_check = xlsx_src.get_columns(row_check_index)
                id_check = self._get_col_id(row_check)
                if id_check == id_val:
                    count += 1
                    skip_index.append(row_check_index)

            # add to log if more than one entry
            if count > 1:
                id_val.append(count)
                xlsx_log.append_row(id_val)

        xlsx_log.save()
        return True

    def _get_col_id(self, row :[]) -> []:
        return [row[i] for i in self.m_cols]



