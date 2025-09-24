from pathlib import Path
import sys


PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def pytest_addoption(parser):
    parser.addoption("--cov", action="append", default=[], help="stub coverage option")
    parser.addoption("--cov-report", action="append", default=[], help="stub coverage option")
