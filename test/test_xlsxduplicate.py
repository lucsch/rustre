#!/usr/bin/env python3

import os

import pytest
import pandas as pd

from .context import rustre
from .context import UNIT_TEST_PATH_OUTPUT
from .context import UNIT_TEST_PATH


@pytest.fixture
def get_test_file():
    return os.path.join(UNIT_TEST_PATH, "test_duplicate.xlsx")

@pytest.fixture
def get_autoclean_file():
    return os.path.join(UNIT_TEST_PATH, "test_autoclean.xlsx")


def test_duplicate(get_test_file):
    xdupli = rustre.xlsxduplicate.XlsxDuplicate(get_test_file,
                                                os.path.join(UNIT_TEST_PATH_OUTPUT, "log_duplicate.xlsx"),
                                                [0, 1, 3])
    assert xdupli.check_duplicate()


def test_duplicate_str(get_test_file):
    xdupli = rustre.xlsxduplicate.XlsxDuplicate(get_test_file,
                                                os.path.join(UNIT_TEST_PATH_OUTPUT, "log_duplicate_str.xlsx"),
                                                ["0", "1", "3"])
    assert xdupli.check_duplicate()


def test_autoclean_get_list(get_test_file):
    clean_obj = rustre.xlsxduplicate.XlsxAutoClean(get_test_file, [0, 1, 3])
    assert clean_obj.get_columns_names() == ['First Name', 'Last Name', 'Gender', 'Age', 'Email', 'Phone', 'Occupation', 'Custom column 1']

def test_clean(get_test_file):
    clean_obj = rustre.xlsxduplicate.XlsxAutoClean(get_test_file, [0, 1, 3])

    # Clean the DataFrame
    out_filename_str = os.path.join(UNIT_TEST_PATH_OUTPUT, "cleaned_file.xlsx")
    number_rows_before_cleaning = clean_obj.df.shape[0]
    result = clean_obj.clean(order_column_index=-1, out_filename=out_filename_str, ascending=True)
    assert result
    assert clean_obj.df.shape[0] < number_rows_before_cleaning

    # open out_filename_str and read the first line in a panda dataframe
    df_out = pd.read_excel(out_filename_str, header=0)
    #print (df_out)

    ## drop duplicates manually to compare to autoclean
    df_origin = pd.read_excel(get_test_file, header=0)
    df_droped = df_origin.drop([5,7], axis='rows').reset_index(drop=True)
    #print(df_droped)

    assert df_out.equals(df_droped)

def test_clean_with_ordering(get_autoclean_file):
    clean_obj = rustre.xlsxduplicate.XlsxAutoClean(get_autoclean_file, [0, 1, 3])

    # Clean the DataFrame
    out_filename_str = os.path.join(UNIT_TEST_PATH_OUTPUT, "cleaned_file2.xlsx")
    number_rows_before_cleaning = clean_obj.df.shape[0]
    # Assuming df is your DataFrame
    order_index = clean_obj.df.columns.get_loc('Keep')
    assert order_index == 4
    result = clean_obj.clean(order_column_index=order_index, out_filename=out_filename_str, ascending=True)
    assert result
    assert clean_obj.df.shape[0] < number_rows_before_cleaning
    #print(clean_obj.df)
    assert clean_obj.df['Keep'].nunique() == 1

