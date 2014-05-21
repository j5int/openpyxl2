# Fixtures (pre-configured objects) for tests
import pytest

# objects under test


@pytest.fixture
def Image():
    """Image class"""
    from openpyxl2.drawing import Image
    return Image


# utility fixtures

@pytest.fixture
def ws(Workbook):
    """Empty worksheet titled 'data'"""
    wb = Workbook()
    ws = wb.get_active_sheet()
    ws.title = 'data'
    return ws


@pytest.fixture
def datadir():
    """DATADIR as a LocalPath"""
    import os
    from py._path.local import LocalPath
    here = os.path.split(__file__)[0]
    DATADIR = os.path.join(here, "test_data")
    return LocalPath(DATADIR)
