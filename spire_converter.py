# never use xlwings

from spire.xls import *
from spire.xls.common import *


if __name__ == "__main__":

    workbook = Workbook()
    print("please enter the path to your excel file")
    excel_file_path = input(" -> ")
    workbook.LoadFromFile(excel_file_path)

    sheet = workbook.Worksheets[0]

    option = HTMLOptions()
    option.ImageEmbedded = True

    print("please enter the path to your html file")
    html_file_path = input(" -> ")
    sheet.SaveToHtml(html_file_path)

    workbook.Dispose