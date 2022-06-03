#!/usr/bin/env python3

import os

import pytest

from .context import rustre
from .context import UNIT_TEST_PATH_OUTPUT
from .context import UNIT_TEST_PATH

# skipping test if local directory didn't exist
if not os.path.exists(os.path.join(UNIT_TEST_PATH, "_local")):
    pytest.skip("Skipping local test", allow_module_level=True)


@pytest.fixture
def get_test_files():
    return [os.path.join(UNIT_TEST_PATH, "_local", "merge_ordi1.xlsx"),
            os.path.join(UNIT_TEST_PATH, "_local", "merge_ordi2.xlsx"),
            os.path.join(UNIT_TEST_PATH, "_local", "merge_ordi3.xlsx")]


def test_merge_file_exists(get_test_files):
    for file in get_test_files:
        assert os.path.exists(file)

@pytest.mark.skip(reason="Very long test")
def test_merge_two_files(get_test_files):
    xmerge = rustre.xlsxmerge.XlsxMerge(get_test_files, sheet_index=0, header_index=7)
    out_file = os.path.join(UNIT_TEST_PATH_OUTPUT, "test_merge_full.xlsx")
    assert xmerge.merge(out_file)
    xlsx = rustre.xlsxfile.XlsxFile(out_file)
    assert xlsx.get_row_count() == 963