"""
银行流水核对工具 v3.2 — 入口
"""
import multiprocessing
import sys
from gui import ReconciliationApp

if __name__ == "__main__":
    multiprocessing.freeze_support()
    if "--check" in sys.argv:
        from matcher import PSUTIL_AVAILABLE

        if not PSUTIL_AVAILABLE:
            print("MEMORY_MONITOR_MISSING")
            raise SystemExit(2)
        print("READY")
        raise SystemExit(0)
    app = ReconciliationApp()
    app.mainloop()
