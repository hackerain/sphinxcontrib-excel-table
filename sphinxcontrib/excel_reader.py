import os

from docutils.parsers.rst import Directive
from docutils.parsers.rst import directives
import docutils.core

import openpyxl
from jinja2 import Environment, FileSystemLoader


def _get_resource_dir(folder):
    current_path = os.path.dirname(__file__)
    resource_path = os.path.join(current_path, folder)
    return resource_path


class ExcelReader(Directive):

    def run(self):
        loader = FileSystemLoader(_get_resource_dir('templates'))
        _env = Environment(loader=loader,
                           keep_trailing_newline=True,
                           trim_blocks=True,
                           lstrip_blocks=True)
        env = self.state.document.settings.env
        width = 600
        height = None

        template = _env.get_template('table.html')
        html = template.render()
        return [docutils.nodes.raw('', html, format='html')]


def setup(app):
    app.add_directive('excel-reader', ExcelReader)
    return {'version': '0.0.1'}
