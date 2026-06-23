from openpyxl import load_workbook
from mailmerge import MailMerge
import qrcode
from io import BytesIO
import tempfile
import os
from PIL import Image


with MailMerge('cover.docx') as document:
    print(document.get_merge_fields())
    # document.merge(Регистр_номер='2222',
    #            field2='Can be used for merging docx documents')
    diploms = []
    # lisst.append({"Регистр_номер":'2222'})
    # lisst.append({"Регистр_номер":'3333'})
    # document.merge_pages(lisst)
    # document.write('output.docx')

    wb = load_workbook('db.xlsx')
    ws = wb["ИП-1 д"]

# Вывод колонок A и C
    for i, row in enumerate(ws.iter_rows(min_row=2, max_row=20, min_col=1, max_col=7, values_only=True)):
        first, middle, last = row[0].split(' ')
        reg = row[1]
        qr = qrcode.make(f"{first} {middle}")
        filename = f'img/qr_{i}.png'
        qr.save(filename)
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
            "Красный": "Нет",
            "file": os.path.abspath(filename)
            })
        # print(f"{fio[0]}")
    document.merge_pages(diploms)
    document.write('output.docx')

    def generate_qr_code(data, size=300):
        """
        Генерирует QR-код и возвращает как bytes
        """
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(data)
        qr.make(fit=True)
        
        # Создаем изображение
        img = qr.make_image(fill_color="black", back_color="white")
        
        # Конвертируем в bytes
        img_bytes = BytesIO()
        img.save(img_bytes, format='PNG')
        img_bytes.seek(0)
    
        return img_bytes.read()
