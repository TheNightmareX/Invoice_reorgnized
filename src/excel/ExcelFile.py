u"""
Created on 2019\u5e7411\u670815\u65e5

@author: \u778c\u7761\u87f2\u5b50
"""
from .Excel03 import Excel03
from .Excel07 import Excel07
import os

def create_excel(sPath):
    if os.path.isfile(sPath):
        return open_excel(sPath)
    else:
        stuff = sPath[-4:]
        stuff = stuff.upper()
        excel = None
        if stuff == '.XLS':
            excel = Excel03().CreateExcel(sPath)
        else:
            if stuff == 'XLSX' or stuff == 'XLSM':
                excel = Excel07().CreateExcel(sPath)
            else:
                raise Exception(u'\u6587\u4ef6\u683c\u5f0f\u9519\u8bef!')
        return excel


def open_excel(sPath):
    if os.path.isfile(sPath):
        stuff = sPath[-4:]
        stuff = stuff.upper()
        excel = None
        if stuff == '.XLS':
            excel = Excel03().OpenExcel(sPath)
        elif stuff == 'XLSX' or stuff == 'XLSM':
            excel = Excel07().OpenExcel(sPath)
        else:
            raise Exception(u'\u6587\u4ef6\u683c\u5f0f\u9519\u8bef!')
    else:
        raise Exception(u'%s\u6587\u4ef6\u4e0d\u5b58\u5728!' % sPath)
    return excel


def save(objExcelWorkBook):
    objExcelWorkBook.Save()


def close_excel(objExcelWorkBook, bSave=True):
    objExcelWorkBook.CloseExcel(bSave=bSave)


def create_sheet(objExcelWorkBook, strSheetName, strWhere='after', bSave=False):
    objExcelWorkBook.CreateSheet(strSheetName, strWhere=strWhere, bSave=bSave)


def get_sheets_name(objExcelWorkBook):
    return objExcelWorkBook.GetSheetsName()


def sheet_rename(objExcelWorkBook, sheet, strNewName, bSave=False):
    objExcelWorkBook.SheetRename(sheet, strNewName, bSave=bSave)


def copy_sheet(objExcelWorkBook, sheet, strNewSheetName, bSave=False):
    objExcelWorkBook.CopySheet(sheet, strNewSheetName, bSave=bSave)


def delete_sheet(objExcelWorkBook, sheet, bSave=False):
    objExcelWorkBook.DeleteSheet(sheet, bSave=bSave)


def active_sheet(objExcelWorkBook, sheet):
    objExcelWorkBook.ActiveSheet(sheet)


def write_cell(objExcelWorkBook, sheet, strCell, data, bSave=False):
    objExcelWorkBook.WriteCell(sheet, strCell, data, bSave=bSave)


def read_cell(objExcelWorkBook, sheet, strCell):
    return objExcelWorkBook.ReadCell(sheet, strCell)


def write_row(objExcelWorkBook, sheet, strCell, data, bSave=False):
    objExcelWorkBook.WriteRow(sheet, strCell, data, bSave=bSave)


def write_column(objExcelWorkBook, sheet, strCell, data, bSave=False):
    objExcelWorkBook.WriteColumn(sheet, strCell, data, bSave=bSave)


def read_row(objExcelWorkBook, sheet, strCell):
    return objExcelWorkBook.ReadRow(sheet, strCell)


def read_column(objExcelWorkBook, sheet, strCell):
    return objExcelWorkBook.ReadColumn(sheet, strCell)


def insert_row(objExcelWorkBook, sheet, strCell, data, bSave=False):
    objExcelWorkBook.InsertRow(sheet, strCell, data, bSave=bSave)


def insert_column(objExcelWorkBook, sheet, strCell, data, bSave=False):
    objExcelWorkBook.InsertColumn(sheet, strCell, data, bSave=bSave)


def merge_range(objExcelWorkBook, sheet, strRange, option=True, bSave=False):
    objExcelWorkBook.MergeRange(sheet, strRange, option=option, bSave=bSave)


def read_range(objExcelWorkBook, sheet, strRange):
    return objExcelWorkBook.ReadRange(sheet, strRange)


def get_rows_count(objExcelWorkBook, sheet):
    return objExcelWorkBook.GetRowsCount(sheet)


def get_columns_count(objExcelWorkBook, sheet):
    return objExcelWorkBook.GetColumnsCount(sheet)


def delete_row(objExcelWorkBook, sheet, strCell, bSave=False):
    objExcelWorkBook.DeleteRow(sheet, strCell, bSave=bSave)


def delete_column(objExcelWorkBook, sheet, strCell, bSave=False):
    objExcelWorkBook.DeleteColumn(sheet, strCell, bSave=bSave)


def insert_image(objExcelWorkBook, sheet, strCell, sFilePath, fWidth, fHeight, bSave=False):
    objExcelWorkBook.InsertImage(sheet, strCell, sFilePath, fWidth, fHeight, bSave=bSave)


def delete_image(objExcelWorkBook, sheet, objPic, bSave=False):
    objExcelWorkBook.DeleteImage(sheet, objPic, bSave=bSave)


def write_range(objExcelWorkBook, sheet, strCell, data, bSave=False):
    objExcelWorkBook.WriteRange(sheet, strCell, data, bSave=bSave)


def clear_range(objExcelWorkBook, sheet, strRange, bClearFormat=True, bSave=False):
    objExcelWorkBook.ClearRange(sheet, strRange, bClearFormat=bClearFormat, bSave=bSave)


def delete_range(objExcelWorkBook, sheet, strRange, bSave=False):
    objExcelWorkBook.DeleteRange(sheet, strRange, bSave=bSave)


def set_row_height(objExcelWorkBook, sheet, strCell, fHeight, bSave=False):
    objExcelWorkBook.SetRowHeight(sheet, strCell, fHeight, bSave=bSave)


def set_column_width(objExcelWorkBook, sheet, strCell, fWidth, bSave=False):
    objExcelWorkBook.SetColumnWidth(sheet, strCell, fWidth, bSave=bSave)


def set_cell_font_color(objExcelWorkBook, sheet, strCell, listColor, bSave=False):
    objExcelWorkBook.SetCellFontColor(sheet, strCell, listColor, bSave=bSave)


def set_range_font_color(objExcelWorkBook, sheet, strRange, listColor, bSave=False):
    objExcelWorkBook.SetRangeFontColor(sheet, strRange, listColor, bSave=bSave)


def set_cell_color(objExcelWorkBook, sheet, strCell, listColor, bSave=False):
    objExcelWorkBook.SetCellColor(sheet, strCell, listColor, bSave=bSave)


def set_range_color(objExcelWorkBook, sheet, strRange, listColor, bSave=False):
    objExcelWorkBook.SetRangeColor(sheet, strRange, listColor, bSave=bSave)


# if __name__ == '__main__':
    # sPath = 'C:\\Users\\Administrator\\Desktop\\aa11.xls'
    # objExcelWorkBook = CreateExcel(sPath)
    # Save(objExcelWorkBook)
    # sheet = 'Sheet1'
    # strCell = 'E8'
    # print(GetSheetsName(objExcelWorkBook))
    # WriteCell(objExcelWorkBook, sheet, strCell, 'data', bSave=True)
    # listColor = [
     # 255, 255, 0]
    # SetCellColor(objExcelWorkBook, sheet, 'B7', [255, 255, 0], bSave=True)
    # listColor = [0, 255, 255]
    # SetCellColor(objExcelWorkBook, sheet, 'F6', [0, 255, 255], bSave=True)
    # listColor = [0, 255, 255]
    # SetCellColor(objExcelWorkBook, sheet, 'E6', [0, 255, 0], bSave=True)
    # CloseExcel(objExcelWorkBook, bSave=True)