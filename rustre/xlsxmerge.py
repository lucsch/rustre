#!/usr/bin/env python3
import logging
from pathlib import Path
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

    def __init__(self, source_files, sheet_index=0, header_index=1):
        """Constructor"""
        self.m_source_files = source_files
        self.m_sheet_index = sheet_index
        self.m_header_index = header_index
        if source_files is None or len(source_files) == 0:
            raise ValueError("No source files")
        for f in source_files:
            if Path(f).suffix != ".xlsx":
                raise ValueError("Wrong extension found")

    def merge(self, result_file):
        """Merge the files into the result file

            :param result_file: the result filename (xlsx file)
            :type result_file: str
            :return: True if the merge is successful, False otherwise
            :rtype: bool
        """
        # compare headers
        xlsx1 = XlsxFile(self.m_source_files[0], self.m_sheet_index, load_in_memory=False)
        header_xlsx1 = xlsx1.get_columns(self.m_header_index)
        for index in range(1, len(self.m_source_files)):
            xlsx2 = XlsxFile(self.m_source_files[index], self.m_sheet_index, load_in_memory=False)
            header_xlsx2 = xlsx2.get_columns(self.m_header_index)
            if header_xlsx1 != header_xlsx2:
                logging.error("Xlsx headers are not equal")
                logging.error("Header1 : {}".format(header_xlsx1))
                logging.error("Header2 : {}".format(header_xlsx2))
                return False

        # merge first file in the new filename
        if not XlsxFile.create_file(result_file):
            return False

        xlsx_result = XlsxFile(result_file)
        xlsx_source1 = XlsxFile(self.m_source_files[0], self.m_sheet_index, load_in_memory=True)
        m_row_tot = xlsx_source1.get_row_count()
        for r in range(self.m_header_index, m_row_tot + 1):
            my_row = xlsx_source1.get_columns(r)
            xlsx_result.append_row(my_row)

        # merge the other files
        for index in range(1, len(self.m_source_files)):
            xlsx_src = XlsxFile(self.m_source_files[index], self.m_sheet_index, load_in_memory=True)
            for r in range(self.m_header_index + 1, xlsx_src.get_row_count() + 1):
                my_row = xlsx_src.get_columns(r)
                xlsx_result.append_row(my_row)
        xlsx_result.save()
        return True
