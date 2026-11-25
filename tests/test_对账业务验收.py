from pathlib import Path
import os
import subprocess
import sys


def test_对账业务验收在独立进程中只使用技能包模块():
    root = Path(__file__).resolve().parents[1]
    env = os.environ.copy()
    env.update(PYTHONUTF8="1", PYTHONIOENCODING="utf-8")
    result = subprocess.run(
        [sys.executable, "-I", "-X", "utf8", "-m", "pytest", "--import-mode=importlib",
         str(root / "tests/_技能业务验收.py"), "-q"],
        cwd=root,
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=180,
    )
    assert result.returncode == 0, result.stdout + result.stderr
