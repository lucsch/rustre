#!/usr/bin/env python3

import os

import pytest

from .context import rustre
from .context import UNIT_TEST_PATH_OUTPUT
from .context import UNIT_TEST_PATH


def test_init_compare_file_not_existing():
    with pytest.raises(ValueError):
        xcomp = rustre.xlsxcompare.XlsxCompare(os.path.join(UNIT_TEST_PATH, "not_exist.xlsx"),
                                               os.path.join(UNIT_TEST_PATH, "test_name1.xlsx"))


@pytest.fixture
def xlsxcompare_obj():
    return rustre.xlsxcompare.XlsxCompare(os.path.join(UNIT_TEST_PATH, "test_name1.xlsx"),
                                          os.path.join(UNIT_TEST_PATH, "test_name2.xlsx"))


def test_compare_columns_ok(xlsxcompare_obj):
    assert xlsxcompare_obj.compare_headers(1, 1) is True





