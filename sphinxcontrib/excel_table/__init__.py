from .excel_table import ExcelTable

def setup(app):
    app.add_directive('excel-table', ExcelTable)
    return {'version': '0.0.1'}
