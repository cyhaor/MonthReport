# -*- coding: utf-8 -*-
import win32com.client


class EasyExcel:
    """A utility to make it easier to get at Excel.  Remembering
    to save the data is your problem, as is  error handling.
    Operates on one workbook at a time."""

    def __init__(self, filename=None):
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        self.xlApp.Visible = 0
        self.xlApp.DisplayAlerts = 0
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

    def save(self, new_filename=None):
        if new_filename:
            self.filename = new_filename
            self.xlBook.SaveAs(new_filename)
        else:
            self.xlBook.Save()

    def close(self):
        self.xlBook.Close(SaveChanges=0)
        self.xlApp.Application.Quit()
        del self.xlApp

    def set_para(self, parameter, values):
        self.xlBook.Names(parameter).RefersToRange.Value = values

    def save_as_pdf(self, output_pdf):
        self.xlBook.ExportAsFixedFormat(0, output_pdf, 1, 1, 1, 1, 6)  # 指定1-6页打印

    def get_cell(self, sheet, row, col):
        """Get value of one cell"""
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col).Value

    def set_cell(self, sheet, row, col, value):
        """set value of one cell"""
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value
        print("set cell:%s", value)

    def get_range(self, sheet, row1, col1, row2, col2):
        """return a 2d array (i.e. tuple of tuples)"""
        sht = self.xlBook.Worksheets(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value

    def add_picture(self, sheet, picture_name, left, top, width, height):
        """Insert a picture in sheet"""
        sht = self.xlBook.Worksheets(sheet)
        sht.Shapes.AddPicture(picture_name, 1, 1, left, top, width, height)

    def cp_sheet(self):
        """copy sheet"""
        sheets = self.xlBook.Worksheets
        sheets(1).Copy(None, sheets(1))

    def insert_row(self, sheet, row):
        sht = self.xlBook.Worksheets(sheet)
        sht.Rows(row).Insert(1)

    def set_range(self, sheet, top_row, left_col, data):
        """insert a 2d array starting at given location.
        Works out the size needed for itself"""
        bottom_row = top_row + len(data) - 1
        right_col = left_col + len(data[0]) - 1
        sht = self.xlBook.Worksheets(sheet)
        # 清空所当前sheet里的有单元格
        # sht.Cells.Clear()

        sht.Range(
            sht.Cells(top_row, left_col),
            sht.Cells(bottom_row, right_col)
        ).Value = data

    def clear_range(self, sheet, top_row, bottom_row, left_col, right_col):
        sht = self.xlBook.Worksheets(sheet)
        sht.Range(
            sht.Cells(top_row, left_col),
            sht.Cells(bottom_row, right_col)
        ).Value = None

    def get_contiguous_range(self, sheet, row, col):
        """Tracks down and across from top left cell until it
        encounters blank cells; returns the non-blank range.
        Looks at first row and column; blanks at bottom or right
        are OK and return None witin the array"""
        sht = self.xlBook.Worksheets(sheet)
        # find the bottom row
        bottom = row
        while sht.Cells(bottom + 1, col).Value not in [None, '']:
            bottom = bottom + 1
        # right column
        right = col
        while sht.Cells(row, right + 1).Value not in [None, '']:
            right = right + 1
        # 设置第一行若干列为粗体
        return sht.Range(sht.Cells(row, col), sht.Cells(bottom, right)).Value


if __name__ == "__main__":
    xls = EasyExcel("C:\\Users\\cyhao\\PycharmProjects\\MonthReportDraft\\test\\testDemoInput.xlsx")
    pdf = "C:\\Users\\cyhao\\PycharmProjects\\MonthReportDraft\\test\\out.pdf"
    xls.set_para("maxErrorModuleNumPercentDevice1", "100%")
    xls.save_as_pdf(pdf)
    xls.save("C:\\Users\\cyhao\\PycharmProjects\\MonthReportDraft\\test\\testDemoOut.xlsx")
    xls.close()
