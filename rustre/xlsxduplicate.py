#!/usr/bin/env python3
import os
import pandas as pd
from rustre.xlsxfile import XlsxFile


class XlsxDuplicate:
    """
    Check a xlsx file for duplicates.
    :param src_filename:
    :param log_filename:
    :param cols:
    """

    def __init__(self, src_filename: str, log_filename: str, cols: [], sheet_index=0, header_index=1):
        self.m_src_filename = src_filename
        self.m_log_filename = log_filename
        # convert cols to int
        self.m_cols = [int(i) for i in cols]
        self.m_sheet_index = sheet_index
        self.m_header_index = header_index

    def is_valid(self) -> bool:
        if not os.path.exists(self.m_src_filename):
            return False
        if self.m_log_filename == "":
            return False
        if len(self.m_cols) == 0:
            return False
        return True

    def check_duplicate(self) -> bool:
        if not self.is_valid():
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
        for row_index in range(self.m_header_index + 1, xlsx_src.get_row_count() + 1):
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

    def _get_col_id(self, row: []) -> []:
        return [row[i] for i in self.m_cols]


class XlsxAutoClean:
    """
    Automatically Clean a xlsx file from duplicates.
    """

    def __init__(self, src_filename: str, cols: [], sheet_index=0, header_index=0):
        """
        Initializes the object with the given source filename, log filename, columns, sheet index, and header index.

        :param src_filename: (str) The source filename.
        :param cols: (list) The list of columns.
        :param sheet_index: (int) The index of the sheet. Default is 0.
        :param header_index: (int) The index of the header. Default is 0.
        """
        # convert cols to int
        self.m_cols = [int(i) for i in cols]
        self.df = pd.read_excel(src_filename, sheet_name=sheet_index, header=header_index)

    def get_columns_names(self) -> []:
        """
        Return a list of column names from the dataframe.
        """
        return self.df.columns.tolist()

    def clean(self, order_column_index: int, out_filename: str, ascending: bool = True) -> bool:
        """
        Clean the DataFrame by sorting it based on the specified column index and removing duplicate rows.

        Args:
            order_column_index (int): The index of the column to use for sorting.
            out_filename (str): The name of the output file to save the cleaned DataFrame to.
            ascending (bool, optional): Whether to sort in ascending order. Defaults to True.

        Returns:
            bool: True if the DataFrame was cleaned and saved successfully, False otherwise.
        """
        if order_column_index != -1:
            self.df = self.df.sort_values(by=self.df.columns[order_column_index])
        keep_rule = 'first' if ascending else 'last'
        column_names = [self.df.columns[i] for i in self.m_cols]
        self.df.drop_duplicates(subset=column_names, keep=keep_rule, inplace=True)
        self.df.to_excel(out_filename, index=False)
        return True
