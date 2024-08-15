from spire.xls import *
from spire.xls.common import *


if __name__ == "__main__":

    workbook = Workbook()
    workbook.LoadFromFile(r"D:\My-Github\pyhton-projects\excel to HTML\Text Format Samples (2).xlsx")

    sheet = workbook.Worksheets[0]

    option = HTMLOptions()
    option.ImageEmbedded = True

    sheet.SaveToHtml(r"D:\My-Github\pyhton-projects\excel to HTML\test.html")

    workbook.Dispose