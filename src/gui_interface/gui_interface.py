import os
import sys

from PyQt4 import QtGui, QtCore
from src.worker.worker import SlotWorker
from PyQt4.QtCore import pyqtSignal, QString

from enum import Enum


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


class Color(Enum):
    White = '#FFFFFF'
    Black = '#000000'
    Green = '#16A765'
    Red = '#FA573C'
    Yellow = '#FFAD46'
    Yellow_light = '#FAD165'
    Grey = '#E1E1E1'
    Blue = '#3F47FF'
    Purple = '#AF9EEF'


class GuiWidget(QtGui.QWidget):

    def __init__(self):
        super(GuiWidget, self).__init__()
        self._summary_file = ""
        self._current_path = os.getcwd()
        self._step = 0
        self._limit = 17
        self.worker_thread = None
        self.msg_box = QtGui.QMessageBox(self)

        self.init_ui()

    def init_ui(self):
        self.label_date_init = QtGui.QLabel()
        self.label_date_init.setText('Initial Date:')

        self.date_init = QtGui.QLineEdit(self)
        self.date_init.setInputMask("00/00/0000")
        self.date_init.setMaximumWidth(65)
        self.date_init.setText("06/03/2020")

        self.label_date_end = QtGui.QLabel()
        self.label_date_end.setText('Final Date:')

        self.date_end = QtGui.QLineEdit(self)
        self.date_end.setInputMask("00/00/0000")
        self.date_end.setMaximumWidth(65)
        self.date_end.setText("10/03/2020")

        self.btn_output = QtGui.QPushButton(' Output file ', self)
        self.btn_output.clicked.connect(self.show_dialog_out)

        self.output_file = QtGui.QLineEdit(self)
        #self.output_file.setReadOnly(True)
        self.output_file.move(130, 20)
        self.output_file.setAlignment(QtCore.Qt.AlignLeft)

        self.bar = QtGui.QProgressBar(self)
        self.bar.setMinimumHeight(30)
        self.bar.setValue(0)
        #self.bar.setMaximum(13)
        #self.change_color(Color.Green)
        self.bar.setAlignment(QtCore.Qt.AlignCenter)

        self.btn_create_summary = QtGui.QPushButton(' Create Summary File ', self)
        self.btn_create_summary.clicked.connect(self.generate_summary)
        self.btn_create_summary.setEnabled(False)
        self.btn_create_summary.setMinimumHeight(30)

        form = QtGui.QFormLayout()
        form.setMargin(0)
        form.setSpacing(5)
        form.addRow(self.label_date_init, self.date_init)
        form.addRow(self.label_date_end, self.date_end)
        form.addRow(self.btn_output, self.output_file)
        form.addRow(self.btn_create_summary)
        form.addRow(self.bar)

        self.setLayout(form)

    def change_color(self, color):
        """
        It is used to change the color of the progress bar.

        Args:
            color(Color): color used for the progress bar.
        """
        style = """
                QProgressBar {{
                    background-color: #E1E1E1;
                    border: 1px solid grey;
                    text-align: center;
                    font-size: 10pt;
                }}""".format(color=color.value)
        self.bar.setStyleSheet(style)

    def show_dialog_out(self):
        self._summary_file = QtGui.QFileDialog.getSaveFileName(self,
                                                               'Save File',
                                                               self._current_path,
                                                               '*.xlsx')
        if self._summary_file != "" and self._summary_file.split('.')[-1] == "xlsx":
            self.output_file.clear()
            self.output_file.setText(self._summary_file)
            self.btn_create_summary.setEnabled(True)

    def generate_summary(self):
        try:
            self.worker_thread = SlotWorker(
                str(self.output_file.text()), init_date=str(self.date_init.text()),
                final_date=str(self.date_end.text()))
            self.worker_thread.processFinished.connect(self.process_finished)
            self.worker_thread.valueChanged.connect(self.update_bar)
            self.worker_thread.intervalChange.connect(self.update_limit)
            self.worker_thread.start()
            #report = ProcessReport(str(self._log_file), str(self._summary_file))
            #report.process()
        except Exception:
            self.msg_box.setIconPixmap(QtGui.QPixmap(resource_path(r"support/error32.png")))
            self.msg_box.setText("Error in the generation of the summary file.")
            self.msg_box.setWindowTitle("Report generator")
            self.msg_box.exec_()

    def process_finished(self):
        self.msg_box.setIconPixmap(QtGui.QPixmap(resource_path(r'support/success32.png')))
        self.msg_box.setText("Summary file has been generated.")
        self.msg_box.setWindowTitle("Report generator")
        self.msg_box.exec_()

    def update_bar(self):
        """
        It is called by the timer to update the status of the progress bar. It should never be
        100% until the end, and it should never be greater than the limit defined.
        """
        if self._step >= self._limit:
            return
        self._step = self._step + 1
        self.bar.setValue(self._step)

    def update_limit(self, limit):
        """
        This method updates the limit in case the progress bar is needed to be reset
        Args:
            limit(int): new max limit
        """
        self._step = 0
        self._limit = limit
        self.bar.setMaximum(self._limit)


class GuiInterface(QtGui.QMainWindow):
    """Main window of the application, it handles the Gui Interfaces and the Menu bar"""

    def __init__(self, parent=None):
        super(GuiInterface, self).__init__(parent)
        self.form_widget = GuiWidget()
        _widget = QtGui.QWidget()
        _layout = QtGui.QVBoxLayout(_widget)
        _layout.addWidget(self.form_widget)
        self.setCentralWidget(_widget)
        self.setGeometry(300, 300, 500, 100)
        self.setWindowTitle('Nomenaments extractor')

        # Create main menu
        self.mainMenu = self.menuBar()
        self.mainMenu.setNativeMenuBar(False)
        self.fileMenu = self.mainMenu.addMenu('File')
        self.helpMenu = self.mainMenu.addMenu('Help')

        # Add exit button
        self.exitButton = QtGui.QAction('Exit', self)
        self.exitButton.setShortcut('Ctrl+Q')
        self.exitButton.setStatusTip('Exit application')
        self.exitButton.triggered.connect(self.close)
        self.fileMenu.addAction(self.exitButton)

        self.helpButton = QtGui.QAction('About', self)
        self.helpButton.setStatusTip('Info about the app')
        self.helpMenu.addAction(self.helpButton)
        self.helpButton.triggered.connect(self.show_about_message)

    def show_about_message(self):
        """About message info"""
        message_box = QtGui.QMessageBox()
        message_box.about(self, 'Nomenaments extractor',
            '<p>Credits:'
            '<br>Adria Guixa - '
            'adriaguixa@gmail.com</p>'
            '<p><b>Version: 0.0.1</b></p>')
