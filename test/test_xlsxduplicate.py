#!/usr/bin/env python3

import os

import pytest

from .context import rustre
from .context import UNIT_TEST_PATH_OUTPUT
from .context import UNIT_TEST_PATH


@pytest.fixture
def get_test_file():
    return os.path.join(UNIT_TEST_PATH, "test_duplicate.xlsx")


def test_duplicate(get_test_file):
    xdupli = rustre.xlsxduplicate.XlsxDuplicate(get_test_file, os.path.join(UNIT_TEST_PATH_OUTPUT, "log_duplicate.xlsx"), [1,2,4])
    assert xdupli.check_duplicate()

