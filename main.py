from openpyxl import load_workbook
from mailmerge import MailMerge, MailMergeOptions, OptionAutoUpdateFields, OptionKeepFields
import qrcode
from io import BytesIO
import tempfile
import os
from PIL import Image
from pathlib import Path
import warnings
from docx import Document
from docx.oxml.ns import qn
from lxml import etree
import copy

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def getQrPath(first, middle, i):
    qr = qrcode.make(f"{first} {middle}")
    filename = f'img/qr_{i}.png'
    qr.save(filename)
    return Path(filename).absolute().as_posix()


mailmerge_options = MailMergeOptions(
    remove_empty_tables=False,
    auto_update_fields_on_open=OptionAutoUpdateFields.ALWAYS,
    # keep_fields=OptionKeepFields.NONE,
    # merge_if_fields=True,
    table_rows_replace_mode=False)

with MailMerge('cover.docx', options=mailmerge_options) as document:
    # print(document.get_merge_fields())

    diploms = []
    wb = load_workbook('db.xlsx')
    ws_data = wb["ИП-1 д"]
    ws_mark = wb["ИП-1 о"]

    subjects = ""
    hours = ""
    separator = ""
    line_breaks = []


    for i, row in enumerate(ws_mark['A3:B71']):
        s, h = row
        separator = "*"
        subject = str(s.value)
        hour = str(h.value)
        
        subjects += subject + separator
       
        if len(subject) > 90: 
            separator = "**"
            line_breaks.append(i)
        
        hours += hour + separator

# Вывод колонок A и C
    for i, row in enumerate(ws_data.iter_rows(min_row=2, max_row=50, min_col=1, max_col=7, values_only=True)):

        fio = row[0]
        if (fio == None):
            break

        first, middle, last = fio.split(' ')
        red = row[6]
        reg = row[1]
        file = getQrPath(first, middle, i)

        col = i+3
        marks = ""
        for i, col in enumerate(ws_mark.iter_cols(min_col=0, max_col=col, min_row=3, max_row=100, values_only=True)):
            for i, mark in enumerate(col):
                if mark == None:
                    continue
                if i in line_breaks: 
                    separator = "**"
                else:
                    separator = "*"
                marks += str(mark) + separator

        diploms.append({
            "Фамилия": first,
            "Имя": middle,
            "Отчество": last,
            "Регистр_номер": reg,
            "Решение_госуд_Экз_комисии": "17 июня 2026 года",
            "Председатель": "Шомахов Замир Валерьевич",
            "Квалификация": "Программист",
            "Дата_выдачи": "01 июля 2026 года",
            "Специальность": "09.02.07 Информационные системы и программирование",
            "Красный": red,
            "file": file,
            "Дисциплины": subjects,
            "Часы": hours,
            "оценки": marks
        })

    document.merge_templates(diploms, separator='page_break')
    document.write('output.docx')

    doc = Document('output.docx')


# doc = Document('output.docx')

# for paragraph in doc.paragraphs:
#     for run in paragraph.runs:
#         # Находим все разрывы страниц
#         for br in run._element.findall(qn('w:br')):
#             if br.get(qn('w:type')) == 'page':
#                 run._element.remove(br)

# doc.save('output.docx')
