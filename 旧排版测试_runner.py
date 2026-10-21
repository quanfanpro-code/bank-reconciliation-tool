# -*- coding: utf-8 -*-
"""临时：以旧排版路径跑指定测试（Task 8 区分新旧失败用，用完即删）。"""
import os
import sys

os.environ["银行流水核对_旧排版"] = "1"

import pytest

sys.exit(pytest.main(["-q", "--tb=no"] + sys.argv[1:]))
