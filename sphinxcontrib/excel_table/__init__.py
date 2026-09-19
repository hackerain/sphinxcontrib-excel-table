from .excel_table import ExcelTable

def setup(app):
    app.add_directive('excel-table', ExcelTable)
    app.add_css_file('handsontable.full.min.css')
    app.add_js_file('handsontable.full.min.js')
    return {'version': '0.0.1'}
