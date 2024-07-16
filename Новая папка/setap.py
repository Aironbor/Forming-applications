
from setuptools import setup
import platform
from glob import glob

SETUP_DICT = {
    'name': 'Заявки в производство',
    'version': '1.0',
    'description': 'Формирование заявок для производства',
    'author': 'Ivan Metliaev',
    'author_email': 'ivan.metliaev.helper@gmail.com',
    'data_files': [
        ('', glob(r'C:\Windows\SYSTEM32\msvcp100.dll')),
        ('platforms', glob(r'C:\Users\IvanW\AppData\Local\Programs\Python\Python39\Lib\site-packages\PyQt5\Qt5\plugins\platforms\qwindows.dll')),
        ('sqldrivers', glob(r'C:\Users\IvanW\AppData\Local\Programs\Python\Python39\Lib\site-packages\PyQt5\Qt5\plugins\sqldrivers\qsqlite.dll')),
        ('qtcoredrivers', glob(r'C:\Users\IvanW\AppData\Local\Programs\Python\Python39\Lib\site-packages\PyQt5\Qt5\bin\Qt5Core.dll')),
        ('qtguidrivers', glob(r'C:\Users\IvanW\AppData\Local\Programs\Python\Python39\Lib\site-packages\PyQt5\Qt5\bin\Qt5Gui.dll')),
        ('qtwidgetdrivers', glob(r'C:\Users\IvanW\AppData\Local\Programs\Python\Python39\Lib\site-packages\PyQt5\Qt5\bin\Qt5Widgets.dll')),
    ],
    'windows': [{'script': 'main.py'}],
    'options': {
        'py2exe': {
            'includes': [
                "lxml._elementpath", "PyQt5.QtCore", "PyQt5.QtGui", "PyQt5.QtWidgets", "images_store",
                "message_widgets", "PyPDF2", "docx2pdf", "PIL", "docx", "subprocess", "logging"],
        },
    }
}

if platform.system() == 'Windows':
    import py2exe
    SETUP_DICT['windows'] = [{
        'Name': 'Ivan Metliaev',
        'product_name': 'Формирование заявок для производства',
        'version': '1.0',
        'description': 'Формирование заявок для производства',
        'copyright': '© 2024, ivan.metliaev.helper@gmail.com. All Rights Reserved',
        'script': 'main.py',
        'icon_resources': [(0, r'applications.ico')]
    }]
    SETUP_DICT['zipfile'] = None

setup(**SETUP_DICT)