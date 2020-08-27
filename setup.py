from setuptools import setup

APP = ['CMRBComp.py']
DATA_FILES = []
PKGS = [
    #'xlsxwriter',
    #'xlrd',
]

INCLDS = [
    'sys',
    'os',
    #'xlsxwriter',
    #'xlrd',
    'datetime',
    'tkinter',
]

OPTIONS = {
    'iconfile': 'icons/CM-RB-Comp-Logo-L',
    'argv_emulation':True,
    'packages': PKGS,
    'includes': INCLDS,
}

setup(
    app = APP,
    data_files = DATA_FILES,
    options = {'py2app': OPTIONS},
    setup_requires = ['py2app'],
)