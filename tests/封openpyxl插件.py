# -*- coding: utf-8 -*-
"""pytest 插件：全程封锁 openpyxl，验证测试套件零依赖。"""
import sys


class _封:
    def find_module(self, name, path=None):
        return self if name == "openpyxl" or name.startswith("openpyxl.") else None

    def load_module(self, name):
        raise ImportError("openpyxl 已封锁: " + name)


def pytest_configure(config):
    sys.meta_path.insert(0, _封())
