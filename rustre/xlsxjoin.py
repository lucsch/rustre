#!/usr/bin/env python3
import os.path

import pandas as pd
import wx


class XlsxJoin:
    """Join two xlsx files based on a column"""

    def __init__(self, base_xlsx: str, base_header: int, base_sheet: int = 0):
        self.m_base_xlsx_file = base_xlsx
        self.m_base_header = base_header
        self.m_base_sheet = base_sheet

    def join(self, second_file: str, second_header: int, second_sheet: int, base_column: int, second_col: int,
             out_file: str):
        if not os.path.exists(self.m_base_xlsx_file):
            wx.LogError("{} file didn't exists!".format(self.m_base_xlsx_file))
            return False

        if not os.path.exists(second_file):
            wx.LogError("{} file didn't exists!".format(second_file))
            return False

        df1 = self._get_dataframe(self.m_base_xlsx_file, self.m_base_sheet, self.m_base_header)
        df2 = self._get_dataframe(second_file, second_sheet, second_header)
        df_merge = pd.merge(df1, df2, on=[df1.columns[base_column], df2.columns[second_col]], how="outer")
        df_merge.to_excel(out_file)
        return True

    def _get_dataframe(self, xlsx_file: str, no_sheet: int, no_header: int) -> pd.DataFrame:
        df = pd.read_excel(xlsx_file, sheet_name=no_sheet, header=no_header)
        return df
