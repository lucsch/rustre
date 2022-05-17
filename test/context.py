#!/usr/bin/env python3
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import rustre  # noqa: F401,E402

UNIT_TEST_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "data"))
UNIT_TEST_PATH_OUTPUT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "unit_test_output"))

