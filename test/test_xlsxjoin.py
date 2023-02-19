#!/usr/bin/env python3

import os

import pytest

from .context import rustre
from .context import UNIT_TEST_PATH_OUTPUT
from .context import UNIT_TEST_PATH


@pytest.fixture()
def get_base_filename():
    return os.path.join(UNIT_TEST_PATH, "test_join1.xlsx")


@pytest.fixture()
def get_second_filename():
    return os.path.join(UNIT_TEST_PATH, "test_join2.xlsx")


def test_join_xlsx(get_base_filename, get_second_filename):
    outfile = os.path.join(UNIT_TEST_PATH_OUTPUT, "test_xlsxjoin.xlsx")

    xjoin = rustre.XlsxJoin(get_base_filename, base_sheet=0, base_header=2)
    assert xjoin.join(get_second_filename, second_header=0, second_sheet=0, base_column=1, second_col=0, out_file=outfile)


def test_join_failed_xlsx(get_base_filename, get_second_filename):
    xjoin = rustre.XlsxJoin(get_base_filename, base_sheet=0, base_header=2)
    assert xjoin.join(os.path.join(UNIT_TEST_PATH, "wrong_file.xlsx"), second_header=0, second_sheet=0, base_column=1, second_col=0, out_file="") is False
