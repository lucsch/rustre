#!/usr/bin/env python3
import logging
import os.path

import pytest

from .context import UNIT_TEST_PATH_OUTPUT


@pytest.fixture(scope="session", autouse=True)
def create_test_output_folder(request):
    logging.info("Creating unit test folder: '{}'".format(UNIT_TEST_PATH_OUTPUT))
    if not os.path.exists(UNIT_TEST_PATH_OUTPUT):
        os.mkdir(UNIT_TEST_PATH_OUTPUT)