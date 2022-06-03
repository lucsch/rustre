#!/usr/bin/env python3

import os
import shutil

import pytest

from .context import rustre
from .context import UNIT_TEST_PATH_OUTPUT
from .context import UNIT_TEST_PATH


@pytest.fixture(scope="module")
def get_src_filename():
    src_filename = os.path.join(UNIT_TEST_PATH_OUTPUT, "test_compare_src.xlsx")
    shutil.copyfile(os.path.join(UNIT_TEST_PATH, "test_compare_src.xlsx"), src_filename)
    return src_filename


@pytest.fixture(scope="module")
def get_target_filename():
    return os.path.join(UNIT_TEST_PATH, "test_compare_target.xlsx")


@pytest.fixture(scope="module")
def get_config_file():
    return os.path.join(UNIT_TEST_PATH, "test_compare.ini")


def test_file_exists(get_src_filename):
    assert os.path.exists(get_src_filename)


def test_compare(get_src_filename, get_config_file, get_target_filename):
    xcomp = rustre.xlsxcompare.XlsxCompare(get_config_file, get_src_filename, get_target_filename)
    assert xcomp.do_compare(os.path.join(UNIT_TEST_PATH_OUTPUT, "compare_log1.xlsx"))
    x_src = rustre.xlsxfile.XlsxFile(get_src_filename)
    my_values = x_src.get_columns(2)
    assert my_values[8] == "no"


@pytest.fixture(scope="module")
def get_src_filename2():
    src_filename = os.path.join(UNIT_TEST_PATH_OUTPUT, "test_src.xlsx")
    shutil.copyfile(os.path.join(UNIT_TEST_PATH, "test_src.xlsx"), src_filename)
    return src_filename


@pytest.fixture(scope="module")
def get_target_filename2():
    return os.path.join(UNIT_TEST_PATH, "test_target.xlsx")


@pytest.fixture(scope="module")
def get_config_file2():
    return os.path.join(UNIT_TEST_PATH, "test_compare2.ini")


def test_compare2(get_src_filename2, get_config_file2, get_target_filename2):
    xcomp = rustre.xlsxcompare.XlsxCompare(get_config_file2, get_src_filename2, get_target_filename2)
    assert xcomp.do_compare(os.path.join(UNIT_TEST_PATH_OUTPUT, "compare_log2.xlsx"))
    x_src = rustre.xlsxfile.XlsxFile(get_src_filename2)
    my_values = x_src.get_columns(8)
    assert my_values[7] == "Wife of Alexander Baker"






