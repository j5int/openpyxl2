import pytest


@pytest.mark.parametrize("value, result",
                         [
                          ('s', 's'),
                          (2.0/3, '0.6666666666666666'),
                          (1, '1'),
                          (None, 'None')
                         ]
                         )
def test_safe_string(value, result):
    from openpyxl2.writer.charts import safe_string
    assert safe_string(value) == result
    v = safe_string('s')
    assert v == 's'
