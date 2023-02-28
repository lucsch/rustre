#!/usr/bin/env python3

import os
import shutil

import pytest

from .context import rustre
from .context import UNIT_TEST_PATH_OUTPUT
from .context import UNIT_TEST_PATH

if not os.path.exists(os.path.join(UNIT_TEST_PATH, "_local")):
    pytest.skip("Skipping local test", allow_module_level=True)


@pytest.fixture()
def get_name():
    return os.path.join(UNIT_TEST_PATH, "_local", "join_name.xlsx")


@pytest.fixture()
def get_plate():
    return os.path.join(UNIT_TEST_PATH, "_local", "join_plates.xlsx")


def test_join_xlsx_local(get_name, get_plate):
    outfile = os.path.join(UNIT_TEST_PATH_OUTPUT, "test_xlsxjoin_local.xlsx")

    xjoin = rustre.XlsxJoin(get_name, base_sheet=0, base_header=0)
    assert xjoin.join(get_plate, second_header=0, second_sheet=0, base_column=0, second_col=20, out_file=outfile)

