# this one doesn't work go to the other file and specify your paths

import xlwings as xw

def get_text_format(sheet, cell):
    if not cell.value:
        return ""

    text = cell.value
    if not text:
        return ""

    html_parts = []
    cell_length = len(text)
    print("..")
    
    # To handle mixed formatting within a cell
    for i in range(cell_length):
        print(".....")
        char = text[i]
        print("...")
        char_range = sheet.characters(i + 1, 1)
        print("..")
        char_format = char_range.Font

        # Detecting formatting for this character
        bold = char_format.Bold
        italic = char_format.Italic
        underline = char_format.Underline
        superscript = char_format.Superscript
        subscript = char_format.Subscript
        
        # Converting Excel formatting values to boolean
        bold = bold != 0
        italic = italic != 0
        underline = underline != 0
        superscript = superscript != 0
        subscript = subscript != 0
        
        # Create HTML parts based on detected formatting
        word_html = char
        if bold:
            word_html = f"<strong>{word_html}</strong>"
        if italic:
            word_html = f"<em>{word_html}</em>"
        if underline:
            word_html = f"<u>{word_html}</u>"
        if superscript:
            word_html = f"<sup>{word_html}</sup>"
        if subscript:
            word_html = f"<sub>{word_html}</sub>"
        
        html_parts.append(word_html)

    return "".join(html_parts)

def get_column_text(sheet, column_letter):
    last_row = sheet.range(f'{column_letter}1').end("down").row
    column_range = sheet.range(f'{column_letter}1:{column_letter}{last_row}')
    column_data = []
    for cell in column_range:
        if cell.value:  # Only process non-empty cells
            print(f"Processing cell: {cell.address}, Value: {cell.value}")
            text = get_text_format(sheet, cell)
            column_data.append(text)
    return column_data

excel = xw.App(visible=False)
workbook = excel.books.open(r'D:\My-Github\pyhton-projects\excel to HTML\Text Format Samples (2).xlsx')
sheet = workbook.sheets['Sheet1']
column_letter = 'A'

column_text = get_column_text(sheet, column_letter)

html_file = r'D:\My-Github\pyhton-projects\excel to HTML\test.html'
with open(html_file, 'w', encoding="utf-8") as file:
    file.write('<html>\n')
    file.write('<body>\n')
    for text in column_text:
        file.write(f'<p>{text}</p>\n')
    file.write('</body>\n')
    file.write('</html>\n')

print("Content printed!")

workbook.close()
excel.quit()
