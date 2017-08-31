# sphinxcontrib-excel-table

A sphinx extension to render an excel table as an excel-alike table in Sphinx documentation.

This contrib is inspired by sphinxcontrib-excel, which uses pyexcel and handsontable to do the
work, but the sphinxcontrib-excel-reader will instead use openpyxl and handsontable, mainly to
support the following features:

* merged cell
* display specific sheet
* display specific section in one sheet
* display specific rows in one sheet

## Installation

You can install it via pip:

```
$ pip install sphinxcontrib-excel-table
```

or install it from source code:

```
$ git clone https://github.com/hackerain/sphinxcontrib-excel-table.git
$ cd sphinxcontrib-excel-table
$ python setup.py install
```
