#!/usr/bin/env python3

import os

import pytest

from .context import rustre
from .context import UNIT_TEST_PATH_OUTPUT
from .context import UNIT_TEST_PATH


@pytest.fixture
def get_test_files():
    return [os.path.join(UNIT_TEST_PATH, "test_name1.xlsx"), os.path.join(UNIT_TEST_PATH, "test_name2.xlsx")]


def test_merge_two_files(get_test_files):
    xmerge = rustre.xlsxmerge.XlsxMerge(get_test_files)
    assert xmerge.merge(os.path.join(UNIT_TEST_PATH_OUTPUT, "test_merge1.xlsx"))
