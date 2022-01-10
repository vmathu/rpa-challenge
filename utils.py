from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem

import os


class XlsxSaver:

    def __init__(self, worksheet: str, path: str):
        self.excel = Files()
        self.lib = FileSystem()
        dir = f'{os.getcwd()}/{path}'
        if self.lib.does_file_exist(dir):
            self.excel.open_workbook(path)
        else:
            self.excel.create_workbook(path=path)
        self.excel.create_worksheet(
            name=worksheet,
            exist_ok=True
        )
        self._path = path

    def _fill_headers(self, col, headers):
        for header in headers:
            self.excel.set_cell_value(1, col, header)
            col += 1

    def _fill_row(self, row_num, col, value):
        self.excel.set_cell_value(row_num, col, value)

    def _fill_rows(self, rows):
        self.excel.append_rows_to_worksheet(rows)

    def _save_workbook(self):
        self.excel.save_workbook(self._path)

    def _close(self):
        self.excel.close_workbook()
