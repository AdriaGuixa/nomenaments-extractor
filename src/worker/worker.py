import xlsxwriter
import requests
import datetime
from bs4 import BeautifulSoup

from PyQt4 import QtCore
from PyQt4.QtCore import QThread, pyqtSignal


class ReportCreator(object):

    def __init__(self, output_file, url='https://antiga.sindicat.net/nomenaments/avui/?data='):
        self._output_file = output_file
        self._url = url
        self._workbook = xlsxwriter.Workbook(self._output_file)

    def get_table(self, date):
        return_table = []
        url = self._url + str(date)
        res = requests.get(url)
        if "charset" in res.headers.get("content-type", "").lower():
            encoding = res.encoding
        else:
            encoding = None
        soup = BeautifulSoup(res.content, 'html.parser', from_encoding=encoding)
        table = soup.find('table')
        table_rows = table.find_all('tr')
        iter_table_rows = iter(table_rows)
        next(iter_table_rows)
        for tr in iter_table_rows:
            td = tr.find_all('td')
            row = [i.text for i in td]
            if row[0] == '17' or row[0] == '3':
                return_table.append(row)
        return return_table

    @staticmethod
    def validate_date(date):
        date_format = '%d/%m/%Y'
        try:
            date_obj = datetime.datetime.strptime(date, date_format)
            return date_obj.strftime('%d/%m/%Y')
        except ValueError:
            raise ValueError("Incorrect data format, should be DD/MM/YYYY")

    @staticmethod
    def get_all_dates(initial_date, final_date):
        all_dates = []
        date_format = '%d/%m/%Y'
        init_date = datetime.datetime.strptime(str(initial_date), date_format)
        fin_date = datetime.datetime.strptime(str(final_date), date_format)
        end_date = True
        valid_date = init_date
        while end_date:
            all_dates.append(init_date.strftime('%d/%m/%Y'))
            if init_date == fin_date:
                end_date = False
            init_date += datetime.timedelta(days=1)

        return all_dates

    def create_table(self, table_list):
        worksheet = self._workbook.add_worksheet('Nomenaments')
        # Start from the first cell. Rows and columns are zero indexed.
        row = 0

        worksheet.write(row, 0, 'SERVEI TERRITORIAL')
        worksheet.write(row, 1, 'DATA')
        worksheet.write(row, 2, 'NUMERO')
        worksheet.write(row, 3, 'CENTRE')
        worksheet.write(row, 4, 'JORNADA')
        worksheet.write(row, 5, 'INICI')
        worksheet.write(row, 6, 'FINAL')
        worksheet.write(row, 7, 'PERFIL')
        worksheet.write(row, 8, 'PROC')
        row += 1

        # Iterate over the data and write it out row by row.
        for item in table_list:
            for counter, value_item in enumerate(item):
                worksheet.write(row, counter, value_item)
            row += 1
        self._workbook.close()


class SlotWorker(QtCore.QThread):
    intervalChange = pyqtSignal(int)
    valueChanged = pyqtSignal()
    processFinished = pyqtSignal()

    def __init__(self, output_file, init_date='01/01/2020', final_date='19/03/2020'):
        super(QThread, self).__init__()
        self.counter = 0
        self._init_date = init_date
        self._final_date = final_date
        self._report_creator = ReportCreator(output_file=output_file)

    def run(self):
        all_dates = []
        all_dates = self._report_creator.get_all_dates(self._init_date, self._final_date)
        self._update_limit(len(all_dates)+1)
        all_tables = []
        for item in all_dates:
            all_tables += self._report_creator.get_table(item)
            self._update_counter()

        self._report_creator.create_table(all_tables)
        self._update_counter()
        self.processFinished.emit()


    def _update_counter(self):
        """Function used to inform slot_interface object to update the current counter value"""
        self.valueChanged.emit()

    def _update_limit(self, limit):
        """Function used to inform slot_interface object to update the limit of the counter"""
        self.intervalChange.emit(limit)
