#!/usr/bin/env python3

import os
import shutil

import pytest

from .context import rustre
from .context import UNIT_TEST_PATH_OUTPUT
from .context import UNIT_TEST_PATH

if not os.path.exists(os.path.join(UNIT_TEST_PATH, "_local")):
    print("Path exist!!!!!")
    pytest.skip("Skipping local test", allow_module_level=True)

@pytest.fixture(scope="module")
def get_src_local():
    src_filename = os.path.join(UNIT_TEST_PATH_OUTPUT, "bdd_local.xlsx")
    shutil.copyfile(os.path.join(UNIT_TEST_PATH, "_local", "bdd.xlsx"), src_filename)
    return src_filename


@pytest.fixture(scope="module")
def get_target_local():
    return os.path.join(UNIT_TEST_PATH, "_local", "ordi.xlsx")


@pytest.fixture(scope="module")
def get_config_local():
    return os.path.join(UNIT_TEST_PATH, "_local", "model.ini")


def test_file_exists(get_src_local):
    assert os.path.exists(get_src_local)


def test_compare(get_src_local, get_config_local, get_target_local):
    xcomp = rustre.xlsxcompare.XlsxCompare(get_config_local, get_src_local, get_target_local)
    assert xcomp.do_compare(os.path.join(UNIT_TEST_PATH_OUTPUT, "compare_local.xlsx"))
