#!/usr/bin/env python3

import os

import pytest

from .context import rustre
from .context import UNIT_TEST_PATH_OUTPUT
from .context import UNIT_TEST_PATH


def test_init_none():
    with pytest.raises(ValueError):
        rustre.xlsxmerge.XlsxMerge(None)
    with pytest.raises(ValueError):
        rustre.xlsxmerge.XlsxMerge([])


@pytest.fixture
def get_test_files():
    return [os.path.join(UNIT_TEST_PATH, "test_name1.xlsx"), os.path.join(UNIT_TEST_PATH, "test_name2.xlsx")]


def test_merge_two_files(get_test_files):
    xmerge = rustre.xlsxmerge.XlsxMerge(get_test_files)
    out_file = os.path.join(UNIT_TEST_PATH_OUTPUT, "test_merge1.xlsx")
    assert xmerge.merge(out_file)
    xlsx = rustre.xlsxfile.XlsxFile(out_file)
    assert xlsx.get_row_count() == 11


def test_merge_headers_differs(get_test_files):
    xmerge = rustre.xlsxmerge.XlsxMerge([get_test_files[0], os.path.join(UNIT_TEST_PATH, "test_compare_src.xlsx")])
    out_file = os.path.join(UNIT_TEST_PATH_OUTPUT, "test_merge2.xlsx")
    assert xmerge.merge(out_file) is False


def test_merge_output_path_didnt_exist(get_test_files):
    xmerge = rustre.xlsxmerge.XlsxMerge(get_test_files)
    out_file = os.path.join(UNIT_TEST_PATH_OUTPUT, "NOT_EXISTING_PATH", "test_merge2.xlsx")
    assert xmerge.merge(out_file) is False

