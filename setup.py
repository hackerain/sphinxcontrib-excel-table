from setuptools import setup, find_packages
import sys

PY2 = sys.version_info[0] == 2
PY26 = PY2 and sys.version_info[1] < 7

NAME = 'sphinxcontrib-excel-table'
AUTHOR = 'Guangyu Suo'
VERSION = '1.0.8'
EMAIL = 'yugsuo@gmail.com'
LICENSE = 'Apache License'
DESCRIPTION = (
    'Embed excel file as an excel-alike table into sphinx documentation.' +
    ''
)
URL = 'https://github.com/hackerain/sphinxcontrib-excel-table'
DOWNLOAD_URL = '%s/archive/%s.tar.gz' % (URL, VERSION)
FILES = ['README.rst', 'CHANGELOG.rst']
KEYWORDS = [
    'python',
    'sphinx',
    'excel'
]

CLASSIFIERS = [
    'Topic :: Office/Business',
    'Topic :: Utilities',
    'Topic :: Software Development :: Libraries',
    'Programming Language :: Python',
    'Framework :: Sphinx :: Extension',
    'Intended Audience :: Developers',
    'Programming Language :: Python :: 2.6',
    'Programming Language :: Python :: 2.7',
    'Programming Language :: Python :: 3.3',
    'Programming Language :: Python :: 3.4',
    'Programming Language :: Python :: 3.5',
    'Programming Language :: Python :: 3.6',
    'License :: OSI Approved :: Apache Software License',
]

INSTALL_REQUIRES = [
    'docutils',
    'sphinx',
    'openpyxl',
    'jinja2',
]


PACKAGES = find_packages(exclude=['ez_setup', 'examples', 'tests'])
EXTRAS_REQUIRE = {
}

def read_files(*files):
    """Read files into setup"""
    text = ""
    for single_file in files:
        content = read(single_file)
        text = text + content + "\n"
    return text


def read(afile):
    """Read a file into setup"""
    with open(afile, 'r') as opened_file:
        content = filter_out_test_code(opened_file)
        content = "".join(list(content))
        return content


def filter_out_test_code(file_handle):
    found_test_code = False
    for line in file_handle.readlines():
        if line.startswith('.. testcode:'):
            found_test_code = True
            continue
        if found_test_code is True:
            if line.startswith('  '):
                continue
            else:
                empty_line = line.strip()
                if len(empty_line) == 0:
                    continue
                else:
                    found_test_code = False
                    yield line
        else:
            for keyword in ['|version|', '|today|']:
                if keyword in line:
                    break
            else:
                yield line

if __name__ == '__main__':
    setup(
        name=NAME,
        author=AUTHOR,
        version=VERSION,
        author_email=EMAIL,
        description=DESCRIPTION,
        url=URL,
        download_url=DOWNLOAD_URL,
        long_description=read_files(*FILES),
        license=LICENSE,
        keywords=KEYWORDS,
        extras_require=EXTRAS_REQUIRE,
        tests_require=['nose'],
        install_requires=INSTALL_REQUIRES,
        packages=PACKAGES,
        include_package_data=True,
        zip_safe=False,
        classifiers=CLASSIFIERS
    )
