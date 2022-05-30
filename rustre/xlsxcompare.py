#!/usr/bin/env python3
from configparser import ConfigParser
from rustre.xlsxfile import XlsxFile


class Config:
    """Parse and store config values found in the ".ini" file

        :param header: the header group name (either SOURCE or TARGET)
        :type header: str
        :param config_file: the filename of the ini file
        :type config_file: str
    """
    def __init__(self, header, config_file):
        """ Constructor """
        self.conf = ConfigParser()
        self.conf.read(config_file)
        self.conf.sections()
        self.m_header = header
        self.m_id_cols = self._as_list_comma("id_col")
        self.m_skip_col = self.conf.getint(self.m_header, "skip_col", fallback=None)
        self.m_skip_col_values = self._as_list_new_line("skip_col_values")
        self.m_col_compare = self.conf.getint(self.m_header, "col_compare")
        self.m_col_copy = self._as_list_comma("col_copy")

    def _as_list_comma(self, config_value):
        my_list = self.conf.get(self.m_header, config_value, fallback=None)
        if my_list is None:
            return None
        my_list = my_list.split(",")
        if my_list == ['']:
            return None
        return [int(i) for i in my_list]

    def _as_list_new_line(self, config_value):
        my_list = self.conf.get(self.m_header, config_value, fallback=None)
        if my_list is None:
            return None
        return my_list.splitlines()

    def get_row_id(self, row):
        id_row = []
        for index in self.m_id_cols:
            id_row.append(row[index])
        return id_row

    def do_skip_row(self, row):
        # check if target row must be skipped
        if self.m_skip_col is not None:
            for val in self.m_skip_col_values:
                if row[self.m_skip_col] == val:
                    return True
        return False

    def do_col_copy(self, dest_list, col_order, row_data):
        """
        Copy all cols listed in col_copy into dest_list in the order specified by col_index
        :param dest_list: A list for storing the copied cols
        :param col_index: A list of cols order
        :param row_data: The data to reorder
        :return: the modified dest_list
        """
        for index, col_index in enumerate(self.m_col_copy):
            dest_list[col_order[index]] = row_data[col_index]
        return dest_list


class XlsxCompare:
    """Compare two xlsx files

        :param config_file: the filepath of the config file (.ini)
        :type config_file: str
        :param file_source: the xslx source filename
        :type file_source: str
        :param file_target: the xlsx target filename
        :type file_target: str
    """

    def __init__(self, config_file, file_source, file_target):
        """ Constructor"""
        self.m_config_file = config_file
        self.m_file_source = file_source
        self.m_file_target = file_target

        # open the config file
        self.m_conf_src = Config("SOURCE", self.m_config_file)
        self.m_conf_target = Config("TARGET", self.m_config_file)

        # open the files
        self.m_xlsx_src = XlsxFile(self.m_file_source, sheet_number=0)
        self.m_xlsx_target = XlsxFile(self.m_file_target, sheet_number=0)


    def do_compare(self, log_file):
        """Compare source with target and modify source based on the data model defined in Config

            :param log_file: xlsx file for saving a log file
            :type log_file: str
            :return: True or False
            :rtype: bool
        """

        # create result log file
        XlsxFile.create_file(log_file)
        xlsx_result = XlsxFile(log_file)
        result_header = self.m_conf_target.get_row_id(self.m_xlsx_target.get_columns(1))
        result_header.append("STATUS")
        xlsx_result.append_row(result_header)

        # iterate all row in target file
        for target_row_index in range(2, self.m_xlsx_target.get_row_count()+1):
            row_target = self.m_xlsx_target.get_columns(target_row_index)
            # id_target = self._get_id(row_target, self.m_conf_target)
            id_target = self.m_conf_target.get_row_id(row_target)

            # do we need to skip this row ?
            if self.m_conf_target.do_skip_row(row_target):
                row_write = self.m_conf_target.get_row_id(row_target)
                row_write.append("SKIPPED")
                xlsx_result.append_row(row_write)
                continue

            # iterate all row in source file
            row_found = False
            for src_row_index in range(2, self.m_xlsx_src.get_row_count()+1):
                row_src = self.m_xlsx_src.get_columns(src_row_index)
                # id_src = self._get_id(row_src, conf_src)
                id_src = self.m_conf_src.get_row_id(row_src)
                if id_src == id_target:
                    row_found = True

                    # check if row has changed
                    print(row_src[self.m_conf_src.m_col_compare])
                    print(row_target[self.m_conf_target.m_col_compare])
                    if row_src[self.m_conf_src.m_col_compare] != row_target[self.m_conf_target.m_col_compare]:
                        # modify the src
                        self.do_row_change(row_target, src_row_index)
                        # self.m_xlsx_src.change_value(conf_src.m_col_compare+1,
                        #                       src_row_index,
                        #                       row_target[self.m_conf_target.m_col_compare])

                        # add the status to the log
                        row_write = self.m_conf_target.get_row_id(row_target)
                        row_write.append("CHANGED")
                        xlsx_result.append_row(row_write)
                        break

            # target row isn't found in src... add it
            if not row_found:
                # add row to the src
                self.do_row_add(row_target)
                # row_target_formated = self._get_target_formated_row(row_target, self.m_conf_target)
                # self.m_xlsx_src.append_row(row_target_formated)

                # add row to the log
                row_write = self.m_conf_target.get_row_id(row_target)
                row_write.append("ADDED")
                xlsx_result.append_row(row_write)

        xlsx_result.save()
        self.m_xlsx_src.save()
        return True

    def do_row_change(self, row_target, src_index):
        """
        Called when the row has changed
        :param row_target: the actual row value of the target
        :param src_index: the index of the row to modify in the source file
        :return:
        """
        self.m_xlsx_src.change_value(self.m_conf_src.m_col_compare+1, src_index,
                                     row_target[self.m_conf_target.m_col_compare])
        # TODO: Check if we need to modify all the columns or only the compare one...
        # for example if the address has changed

    def do_row_add(self, row_target):
        # create empty list
        my_new_row = [None] * len(self.m_xlsx_src.get_columns(1))
        my_new_row = self.m_conf_target.do_col_copy(my_new_row, self.m_conf_src.m_col_copy, row_target)
        self.m_xlsx_src.append_row(my_new_row)


    def _get_target_formated_row(self, row_target, conf_target):
        order_list = [row_target[i] for i in conf_target.m_col_order]
        return order_list


