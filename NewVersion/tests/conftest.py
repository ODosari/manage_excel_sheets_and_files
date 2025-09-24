from pathlib import Path
import sys


PACKAGE_ROOT = Path(__file__).resolve().parents[1]
if str(PACKAGE_ROOT) not in sys.path:
    sys.path.insert(0, str(PACKAGE_ROOT))


def pytest_addoption(parser):
    parser.addoption("--cov", action="append", default=[], help="stub coverage option")
    parser.addoption("--cov-report", action="append", default=[], help="stub coverage option")
