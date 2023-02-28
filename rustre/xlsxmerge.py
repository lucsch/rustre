#!/usr/bin/env python3
import logging
import pandas as pd
from pathlib import Path

import wx

from rustre.xlsxfile import XlsxFile


class XlsxMerge:
    """Merge multiple xlsx files

        :param source_files: list of xlsx filenames
        :type source_files: list[str]
        :param sheet_index: zero based index of the xlsx sheet (default is 0)
        :type sheet_index: int
        :param header_index: header index (starts at 1) default is 1.
        :type header_index: int
        :raises:
            :ValueError: if source_files is None or empty
    """

    def __init__(self, source_files, sheet_index=0, header_index=0):
        """Constructor"""
        self.m_source_files = source_files
        self.m_sheet_index = sheet_index
        self.m_header_index = header_index
        if source_files is None or len(source_files) == 0:
            raise ValueError("No source files")
        for f in source_files:
            if Path(f).suffix != ".xlsx":
                raise ValueError("Wrong extension found")

    @staticmethod
    def is_headers_equal(base_df: pd.DataFrame, second_df: pd.DataFrame) -> bool:
        diff_cols = base_df.columns.difference(second_df.columns)
        diff_cols2 = second_df.columns.difference(base_df.columns)
        if len(diff_cols) == 0 and len(diff_cols2) == 0:
            return True
        if len(diff_cols) != 0:
            wx.LogError("Headers mismatch in first file, column: {}".format(",".join(diff_cols)))
        if len(diff_cols2) != 0:
            wx.LogError("Headers mismatch in other file, column: {}".format(",".join(diff_cols2)))
        return False

    def merge(self, result_file):
        """Merge the files into the result file

            :param result_file: the result filename (xlsx file)
            :type result_file: str
            :return: True if the merge is successful, False otherwise
            :rtype: bool
        """
        # compare headers and merge
        df = pd.DataFrame()
        df1 = pd.read_excel(self.m_source_files[0], sheet_name=self.m_sheet_index, header=self.m_header_index)
        df = df.append(df1)
        for index in range(1, len(self.m_source_files)):
            df_iter = pd.read_excel(self.m_source_files[index], sheet_name=self.m_sheet_index,
                                    header=self.m_header_index)
            if not self.is_headers_equal(df1, df_iter):
                wx.LogError("Column mismatch")
                return False
            df = df.append(df_iter)

        # try to create the output file
        if not XlsxFile.create_file(result_file):
            return False

        df.to_excel(result_file, index=False)
        return True
