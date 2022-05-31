#!/usr/bin/env python3
from configparser import ConfigParser
from rustre.xlsxfile import XlsxFile


class ColMapping:
    def __init__(self, row):
        self.m_col = None
        self.m_values = []

        if row is not None and row != "":
            my_list = row.split(",")
            self.m_col = int(my_list[0])
            for val in my_list[1:]:
                self.m_values.append(val.replace(";", ","))

    def is_valid_col(self):
        if self.m_col is not None:
            return True
        return False


class ColCondition:
    def __init__(self, row):
        self.m_col_one = -1
        self.m_value = None
        self.m_col_two = -1

        if row is not None and row != "":
            my_list = row.split(",")
            self.m_col_one = int(my_list[0])
            self.m_value = str(my_list[1])
            if self.m_value == "None":
                self.m_value = None
            self.m_col_two = int(my_list[2])

    def is_valid(self):
        if self.m_col_one == -1 and self.m_col_two == -1 and self.m_value is None:
            return False
        return True


class ColStripText:
    def __init__(self, row):
        self.m_col = -1
        self.m_nb_char_to_strip = 0

        if row is not None and row != "":
            my_list = row.split(",")
            self.m_col = int(my_list[0])
            self.m_nb_char_to_strip = int(my_list[1])

    def is_valid(self):
        if self.m_col == -1 and self.m_nb_char_to_strip == 0:
            return False
        return True


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
        self.m_col_join = self._as_list_of_list("col_join")
        self.m_col_mapping = []
        my_mapping_list = self._as_list_new_line("col_mapping")
        if my_mapping_list is not None:
            for row in my_mapping_list:
                col_map = ColMapping(row)
                if col_map.is_valid_col():
                    self.m_col_mapping.append(col_map)
        self.m_col_condition = []
        my_conditions_list = self._as_list_new_line("col_condition")
        if my_conditions_list is not None:
            for col in my_conditions_list:
                col_cond = ColCondition(col)
                if col_cond.is_valid():
                    self.m_col_condition.append(col_cond)

        self.m_col_strip_text = []
        my_strip_list = self._as_list_new_line("col_strip_text")
        if my_strip_list is not None:
            for row in my_strip_list:
                col_strip = ColStripText(row)
                if row is not None:
                    self.m_col_strip_text.append(col_strip)

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

    def _as_list_of_list(self, config_value):
        my_list = self._as_list_new_line(config_value)
        if my_list is not None and "," in my_list[1]:
            my_new_list = []
            for item in my_list:
                my_new_list.append(item.split(","))
            return my_new_list
        return my_list

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

    def do_col_join(self, dest_list, join_cols, row_data):
        if self.m_col_join is None:
            return dest_list

        # iterate source join cols
        for src_index, src_col in enumerate(join_cols):
            # skip empty index
            if src_col == "":
                continue
            # iterate target join cols
            my_joined_value = ""
            my_cols_int = [int(i) for i in self.m_col_join[src_index]]
            for index, col in enumerate(my_cols_int):
                my_joined_value += str(row_data[col]) + " "
            dest_list[int(src_col)] = my_joined_value.rstrip()
        return dest_list

    def do_col_mapping(self, dest_list, col_mapping_obj, row_data):
        if col_mapping_obj is None:
            return dest_list
        for index, map_obj in enumerate(self.m_col_mapping):
            for val_index, val in enumerate(map_obj.m_values):
                if val == str(row_data[map_obj.m_col]):
                    dest_list[col_mapping_obj[index].m_col] = col_mapping_obj[index].m_values[val_index]
        return dest_list

    def do_col_condition(self, dest_list, col_condition_obj, row_data):
        if col_condition_obj is None:
            return dest_list
        for index, cond_obj in enumerate(self.m_col_condition):
            if cond_obj.m_value in row_data[cond_obj.m_col_one]:
                dest_list[col_condition_obj[index].m_col_one] = row_data[cond_obj.m_col_one]
                dest_list[col_condition_obj[index].m_col_two] = col_condition_obj[index].m_value
            else:
                dest_list[col_condition_obj[index].m_col_two] = row_data[cond_obj.m_col_one]
        return dest_list

    def do_col_strip_text(self, dest_list, col_strip_text_obj, row_data):
        if col_strip_text_obj is None:
            return dest_list
        for index, strip_obj in enumerate(self.m_col_strip_text):
            if not strip_obj.is_valid():
                continue
            text_striped = row_data[strip_obj.m_col][strip_obj.m_nb_char_to_strip:]
            dest_list[col_strip_text_obj[index].m_col] = text_striped.strip()
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

                        # add the status to the log
                        row_write = self.m_conf_target.get_row_id(row_target)
                        row_write.append("CHANGED")
                        xlsx_result.append_row(row_write)
                        break

            # target row isn't found in src... add it
            if not row_found:
                # add row to the src
                self.do_row_add(row_target)

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
        my_new_row = self.m_conf_target.do_col_join(my_new_row, self.m_conf_src.m_col_join, row_target)
        my_new_row = self.m_conf_target.do_col_mapping(my_new_row, self.m_conf_src.m_col_mapping, row_target)
        my_new_row = self.m_conf_target.do_col_condition(my_new_row, self.m_conf_src.m_col_condition, row_target)
        my_new_row = self.m_conf_target.do_col_strip_text(my_new_row, self.m_conf_src.m_col_strip_text, row_target)
        print(my_new_row)
        self.m_xlsx_src.append_row(my_new_row)

    def _get_target_formated_row(self, row_target, conf_target):
        order_list = [row_target[i] for i in conf_target.m_col_order]
        return order_list


