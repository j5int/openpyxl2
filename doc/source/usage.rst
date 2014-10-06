Simple usage
============

Write a workbook
----------------
.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl2[.]compat import range
>>> from openpyxl2[.]cell import get_column_letter
>>>
>>> wb = Workbook()
>>>
>>> dest_filename = 'empty_book.xlsx'
>>>
>>> ws = wb.active
>>>
>>> ws.title = "range names"
>>>
>>> for col_idx in range(1, 40):
...     col = get_column_letter(col_idx)
...     for row in range(1, 600):
...         ws.cell('%s%s'%(col, row)).value = '%s%s' % (col, row)
>>>
>>> ws = wb.create_sheet()
>>>
>>> ws.title = 'Pi'
>>>
>>> ws['F5'] = 3.14
>>>
>>> wb.save(filename = dest_filename)


Write a workbook from \*.xltx as \*.xlsx
----------------------------------------
.. ::doctest

>>> from openpyxl import load_workbook
>>>
>>>
>>> wb = load_workbook('sample_book.xltx') #doctest: +SKIP
>>> ws = wb.active #doctest: +SKIP
>>> ws['D2'] = 42 #doctest: +SKIP
>>>
>>> wb.save('sample_book.xlsx') #doctest: +SKIP
>>>
>>> # or you can overwrite the current document template
>>> # wb.save('sample_book.xltx')


Write a workbook from \*.xltm as \*.xlsm
----------------------------------------
.. ::doctest

>>> from openpyxl import load_workbook
>>>
>>>
>>> wb = load_workbook('sample_book.xltm', keep_vba=True) #doctest: +SKIP
>>> ws = wb.active #doctest: +SKIP
>>> ws['D2'] = 42 #doctest: +SKIP
>>>
>>> wb.save('sample_book.xlsm') #doctest: +SKIP
>>>
>>> # or you can overwrite the current document template
>>> # wb.save('sample_book.xltm')


Read an existing workbook
-------------------------
.. :: doctest

>>> from openpyxl import load_workbook
>>> wb = load_workbook(filename = 'empty_book.xlsx')
>>> sheet_ranges = wb['range names']
>>> print(sheet_ranges['D18'].value)
D18


.. note ::

    There are several flags that can be used in load_workbook.

    - `guess_types` will enable or disable (default) type inference when
      reading cells.

    - `data_only` controls whether cells with formulae have either the
      formula (default) or the value stored the last time Excel read the sheet.

    - `keep_vba` controls whether any Visual Basic elements are preserved or
      not (default). If they are preserved they are still not editable.


.. warning ::

    openpyxl does currently not read all possible items in an Excel file so
    images and charts will be lost from existing files if they are opened and
    saved with the same name.


Using number formats
--------------------
.. :: doctest

>>> import datetime
>>> from openpyxl import Workbook
>>> wb = Workbook(guess_types=True)
>>> ws = wb.active
>>> # set date using a Python datetime
>>> ws['A1'] = datetime.datetime(2010, 7, 21)
>>>
>>> ws['A1'].number_format
'yyyy-mm-dd h:mm:ss'
>>>
>>> # set percentage using a string followed by the percent sign
>>> ws['B1'] = '3.14%'
>>>
>>> ws['B1'].value
0.031400000000000004
>>>
>>> ws['B1'].number_format
'0%'


Using formulae
--------------
.. :: doctest

>>> from openpyxl import Workbook
>>> wb = Workbook()
>>> ws = wb.active
>>> # add a simple formula
>>> ws["A1"] = "=SUM(1, 1)"
>>> wb.save("formula.xlsx")

.. warning::
    NB function arguments *must* be separated by commas and not other
    punctuation such as semi-colons


Merge / Unmerge cells
---------------------
.. :: doctest

>>> from openpyxl2[.]workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.merge_cells('A1:B1')
>>> ws.unmerge_cells('A1:B1')
>>>
>>> # or
>>> ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=4)
>>> ws.unmerge_cells(start_row=2,start_column=1,end_row=2,end_column=4)


Inserting an image
-------------------
.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl2[.]drawing import Image
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>> ws['A1'] = 'You should see three logos below'
>>> ws['A2'] = 'Resize the rows and cells to see anchor differences'
>>>
>>> # create image instances
>>> img = Image('logo.png')
>>> img2 = Image('logo.png')
>>> img3 = Image('logo.png')
>>>
>>> # place image relative to top left corner of spreadsheet
>>> img.drawing.top = 100
>>> img.drawing.left = 150
>>>
>>> # the top left offset needed to put the image
>>> # at a specific cell can be automatically calculated
>>> img2.anchor(ws['D12'])
(('D', 12), ('D', 21))
>>>
>>> # one can also position the image relative to the specified cell
>>> # this can be advantageous if the spreadsheet is later resized
>>> # (this might not work as expected in LibreOffice)
>>> img3.anchor(ws['G20'], anchortype='oneCell')
((6, 19), None)
>>>
>>> # afterwards one can still add additional offsets from the cell
>>> img3.drawing.left = 5
>>> img3.drawing.top = 5
>>>
>>> # add to worksheet
>>> ws.add_image(img)
>>> ws.add_image(img2)
>>> ws.add_image(img3)
>>> wb.save('logo.xlsx')


Fold columns (outline)
----------------------
.. :: doctest

>>> import openpyxl2
>>> wb = openpyxl.Workbook(True)
>>> ws = wb.create_sheet()
>>> ws.column_dimensions.group('A','D', hidden=True)
>>> wb.save('group.xlsx')
