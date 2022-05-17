#!/usr/bin/env python3

import os

import pytest

from .context import rustre
from .context import UNIT_TEST_PATH_OUTPUT
from .context import UNIT_TEST_PATH


def test_init_wrong_workbook():
    with pytest.raises(ValueError):
        wrongxlsx = rustre.xlsxfile.XlsxFile(os.path.join(UNIT_TEST_PATH, "not_exist.xlsx"))


def test_init_correct_workbook():
    xlsx = rustre.xlsxfile.XlsxFile(os.path.join(UNIT_TEST_PATH, "test_name1.xlsx"))


@pytest.fixture
def test_xlsx():
    return rustre.xlsxfile.XlsxFile(os.path.join(UNIT_TEST_PATH, "test_name1.xlsx"))


def test_get_columns(test_xlsx):
    my_col = test_xlsx.get_columns()
    assert len(my_col) == 8


def test_create_empty_file():
    test_filename = os.path.join(UNIT_TEST_PATH_OUTPUT, "test_empty.xlsx")
    xlsx = rustre.xlsxfile.XlsxFile.create_file(test_filename)
    assert os.path.exists(test_filename)

