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
from datetime import datetime
import locale

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def format_date_ru(date_str):
    # Парсим исходную дату

    # Словарь с русскими названиями месяцев в родительном падеже
    months = {
        1: 'января', 2: 'февраля', 3: 'марта', 4: 'апреля',
        5: 'мая', 6: 'июня', 7: 'июля', 8: 'августа',
        9: 'сентября', 10: 'октября', 11: 'ноября', 12: 'декабря'
    }

    # Форматируем: день с ведущим нулём, месяц из словаря, год, слово "года"
    return f"{date_str.day:02d} {months[date_str.month]} {date_str.year} года"


def get_srok(attestat, common_data):
    if attestat == "аттестат об основном общем образовании":
        return common_data["srok9"]
    else:
        return common_data["srok11"]


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

    data_vidachi = "01 июля 2026 года"
    gruppa = "ИП-1"
    diploms = []
    wb = load_workbook(r'C:\Users\ganz\OneDrive\Документы\ведомости оценок.xlsx')
    common = load_workbook('utils.xlsx')
    ws_common = common["data"]
    ws_data = wb[gruppa + " д"]
    ws_mark = wb[gruppa + " о"]
    common_data = {}

    for row in ws_common.values:
        grupp_name = row[0]
        if grupp_name == gruppa:
            common_data = {
                "speciality": row[1],
                "reshenie": row[2],
                "predsedatel": row[3],
                "srok9": row[4],
                "srok11": row[5],
                "vidacha": row[6],
                "qualification" : row[7]
            }
            # Вывод колонок A и C
    for i, row in enumerate(ws_data.iter_rows(min_row=2, max_row=50, min_col=1, max_col=7, values_only=True)):
        fio = row[0]
        if (fio == None):
            break

        print(format_date_ru(row[2]))
        first, middle, last = fio.split(' ')
        dt_obj = row[2]
        birthdate = format_date_ru(row[2])
        red = row[6]
        reg = row[1]
        file = getQrPath(first, middle, i)
        attestat_type = str(row[3])
        attestat_year = str(row[4])
        pred_document = attestat_type + ", " + attestat_year + " года"
        srok = get_srok(attestat_type, common_data)
        qualification = common_data["qualification"]
        speciality = common_data["speciality"]
        predsedatel = common_data["predsedatel"]
        reshenie = format_date_ru(common_data["reshenie"])

        subjects = ""
        hours = ""
        marks = ""

        kursovie_subjects = ""
        kursovie_marks = ""

        student_col = i+3
        for mark_col in ws_mark.iter_rows(min_row=3, max_row=100, min_col=0, max_col=student_col, values_only=True):
            subject = mark_col[0]
            hour = mark_col[1]
            mark = mark_col[student_col-1]

            if subject == None:
                break
            if mark == None:
                continue

            subject = str(subject)
            hour = str(hour)
            mark = str(mark)

            separator = "*"

            if ("Курсовая" in subject):
                kursovie_subjects += subject + separator
                kursovie_marks += mark + separator
            else:
                subjects += subject + separator

                if len(subject) >= 90:
                    separator = "**"

                hours += hour + separator
                marks += mark + separator

        diploms.append({
            "Фамилия": first,
            "Имя": middle,
            "Отчество": last,
            "Регистр_номер": reg,
            "число_месяцполностью_год__рождения": birthdate,
            "Решение_госуд_Экз_комисии": reshenie,
            "Председатель": predsedatel,
            "Дата_выдачи": data_vidachi,
            "Специальность": speciality,
            "Красный": red,
            "file": file,
            "Дисциплины": subjects,
            "Часы": hours,
            "оценки": marks,
            "Предыдущий_документ": pred_document,
            "срок": srok,
            "Квалификация": qualification,
            "Специальность": speciality,
            "курсовые": kursovie_subjects,
            "оценка_курсовая": kursovie_marks

        })

    document.merge_templates(diploms, separator='page_break')
    document.write('output.docx')

    # doc = Document('output.docx')


doc = Document('output.docx')

for paragraph in doc.paragraphs:
    for run in paragraph.runs:
        # Находим все разрывы страниц
        for br in run._element.findall(qn('w:br')):
            if br.get(qn('w:type')) == 'page':
                run._element.remove(br)

doc.save('output.docx')
