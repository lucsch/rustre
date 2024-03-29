#!/usr/bin/env python3

import os

import pytest
import pandas as pd
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


def test_merge_three_files(get_test_files):
    list_files = get_test_files
    list_files.append(get_test_files[0])
    assert len(list_files) == 3
    xmerge = rustre.xlsxmerge.XlsxMerge(list_files)
    out_file = os.path.join(UNIT_TEST_PATH_OUTPUT, "test_merge3.xlsx")
    assert xmerge.merge(out_file)
    xlsx = rustre.xlsxfile.XlsxFile(out_file)
    assert xlsx.get_row_count() == 16

def test_merge_headers_differs(get_test_files):
    xmerge = rustre.xlsxmerge.XlsxMerge([get_test_files[0], os.path.join(UNIT_TEST_PATH, "test_compare_src.xlsx")])
    out_file = os.path.join(UNIT_TEST_PATH_OUTPUT, "test_merge2.xlsx")
    assert xmerge.merge(out_file) is False


def test_merge_output_path_didnt_exist(get_test_files):
    xmerge = rustre.xlsxmerge.XlsxMerge(get_test_files)
    out_file = os.path.join(UNIT_TEST_PATH_OUTPUT, "NOT_EXISTING_PATH", "test_merge2.xlsx")
    assert xmerge.merge(out_file) is False


def test_init_failed_wrong_format(get_test_files):
    wrong_file = os.path.join(UNIT_TEST_PATH, "test_compare.ini")
    with pytest.raises(ValueError):
        xmerge = rustre.xlsxmerge.XlsxMerge([get_test_files[0], wrong_file, get_test_files[1]])


def test_headers_equal(get_test_files):
    df1 = pd.read_excel(get_test_files[0], sheet_name=0, header=0)
    df2 = pd.read_excel(get_test_files[1], sheet_name=0, header=0)
    assert rustre.XlsxMerge.is_headers_equal(base_df=df1, second_df=df2)


def test_header_differs(get_test_files):
    df1 = pd.read_excel(get_test_files[0], sheet_name=0, header=0)
    df2 = pd.read_excel(os.path.join(UNIT_TEST_PATH, "test_join1.xlsx"), sheet_name=0, header=2)
    assert rustre.XlsxMerge.is_headers_equal(df1, df2) is False


def test_header_differs2(get_test_files):
    df1 = pd.read_excel(get_test_files[0], sheet_name=0, header=0)
    df2 = pd.read_excel(os.path.join(UNIT_TEST_PATH, "test_compare_src.xlsx"), sheet_name=0, header=0)
    assert rustre.XlsxMerge.is_headers_equal(df1, df2) is False