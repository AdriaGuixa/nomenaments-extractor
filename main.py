import os
import sys
import argparse

from PyQt4 import QtGui

from src.gui_interface import gui_interface

from src.worker import worker


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS,
        # and places our data files in a folder relative to that temp
        # folder named as specified in the datas tuple in the spec file
        base_path = sys._MEIPASS
    except Exception:
        # sys._MEIPASS is not defined, so use the original path
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

if __name__ == '__main__':
    '''
    file_path = r'C:/Projects/nomenaments-app/output.xlsx'
    init_date = '25/08/2019'
    final_date = '14/04/2020'
    all_dates = []
    all_tables = []
    remote_worker = worker.ReportCreator(file_path)
    all_dates = remote_worker.get_all_dates(init_date, final_date)
    for item in all_dates:
        all_tables += remote_worker.get_table(item)

    remote_worker.create_table(all_tables)
    print 'hop'

    '''
    app = QtGui.QApplication(sys.argv)
    # ... the rest of your handling: `sys.exit(app.exec_())`, etc.
    app.setWindowIcon(QtGui.QIcon(resource_path('support/icon_1.png')))
    win = gui_interface.GuiInterface()
    win.show()
    sys.exit(app.exec_())
