"""
银行流水核对工具 v3.0 — 入口
"""
import multiprocessing
from gui import ReconciliationApp

if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = ReconciliationApp()
    app.mainloop()
