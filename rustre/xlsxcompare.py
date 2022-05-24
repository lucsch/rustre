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
        """ Contructor """
        conf = ConfigParser()
        conf.read(config_file)
        conf.sections()
        self.m_id_cols = conf.get(header, "id_col").split(",")
        self.m_id_cols = [int(i) for i in self.m_id_cols]
        self.m_skip_col = conf[header]["skip_col"]
        self.m_skip_col_value = conf[header]["skip_col_value"]
        self.m_col_compare = conf[header]["col_compare"]
        self.m_col_order = conf[header]["col_order"].split(",")
        self.m_col_order = [int(i) for i in self.m_col_order]
        if self.m_skip_col == '':
            self.m_skip_col = None
        else:
            self.m_skip_col = int(self.m_skip_col)
        if self.m_skip_col_value == '':
            self.m_skip_col_value = None
        if self.m_col_compare != '':
            self.m_col_compare = int(self.m_col_compare)


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

    def do_compare(self, log_file):
        """Compare source with target and modify source based on the data model defined in Config

            :param log_file: xlsx file for saving a log file
            :type log_file: str
            :return: True or False
            :rtype: bool
        """
        # open the config file
        conf_src = Config("SOURCE", self.m_config_file)
        conf_target = Config("TARGET", self.m_config_file)
        xlsx_src = XlsxFile(self.m_file_source, sheet_number=0)
        xlsx_target = XlsxFile(self.m_file_target, sheet_number=0)

        # create result log file
        XlsxFile.create_file(log_file)
        xlsx_result = XlsxFile(log_file)
        result_header = self._get_id_row(xlsx_target.get_columns(1), conf_target)
        result_header.append("STATUS")
        xlsx_result.append_row(result_header)

        # iterate all row in target file
        for target_row_index in range(2, xlsx_target.get_row_count()+1):
            row_target = xlsx_target.get_columns(target_row_index)
            id_target = self._get_id(row_target, conf_target)

            # do we need to skip this row ?
            if self._skip_row(row_target, conf_target):
                row_write = self._get_id_row(row_target, conf_target)
                row_write.append("SKIPPED")
                xlsx_result.append_row(row_write)
                continue

            # iterate all row in source file
            row_found = False
            for src_row_index in range(2, xlsx_src.get_row_count()+1):
                row_src = xlsx_src.get_columns(src_row_index)
                id_src = self._get_id(row_src, conf_src)
                if id_src == id_target:
                    row_found = True

                    # check if row has changed
                    if row_src[conf_src.m_col_compare] != row_target[conf_target.m_col_compare]:
                        # modify the src
                        xlsx_src.change_value(conf_src.m_col_compare+1,
                                              src_row_index,
                                              row_target[conf_target.m_col_compare])

                        # add the status to the log
                        row_write = self._get_id_row(row_target, conf_target)
                        row_write.append("CHANGED")
                        xlsx_result.append_row(row_write)
                        break

            # target row isn't found in src... add it
            if not row_found:
                # add row to the src
                row_target_formated = self._get_target_formated_row(row_target, conf_target)
                xlsx_src.append_row(row_target_formated)

                # add row to the log
                row_write = self._get_id_row(row_target, conf_target)
                row_write.append("ADDED")
                xlsx_result.append_row(row_write)

        xlsx_result.save()
        xlsx_src.save()
        return True

    def _skip_row(self, row_target, conf_target):
        # check if target row must be skipped
        if conf_target.m_skip_col is not None:
            if row_target[conf_target.m_skip_col] == conf_target.m_skip_col_value:
                return True
        return False

    def _get_id(self, row, conf):
        my_id = ""
        for col_index in conf.m_id_cols:
            my_id += str(row[col_index])
        return my_id

    def _get_target_formated_row(self, row_target, conf_target):
        order_list = [row_target[i] for i in conf_target.m_col_order]
        return order_list

    def _get_id_row(self, row_target, conf_target):
        id_row = []
        for index in conf_target.m_id_cols:
            id_row.append(row_target[index])
        return id_row
