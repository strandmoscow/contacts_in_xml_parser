import os
import xml.etree.ElementTree as et
import pandas as pd

# Путь к корневой директории с XML файлами
PATH_TO_DIR = 'path/to/dir'

# Выбор названия для выходного файла (без расширения)
OUT_FILE_NAME = 'out_file'

# Выбор расширения для итогового файла между xlsx или csv.
OUT_FILE_EXTENTION = 'csv'  # xlsx или csv


def parse_xml(file_path):
    # Парсинг XML файла
    tree = et.parse(file_path)
    root = tree.getroot()

    # Извлечение номера заявки из пути к файлу
    application_number = os.path.basename(os.path.dirname(file_path))

    # Инициализация переменных для хранения данных
    person_type = ''
    full_name = ''
    org_name = ''
    address = ''
    phone = ''
    email = ''
    first_name_fl = ''
    last_name_fl = ''
    middle_name_fl = ''

    # Поиск необходимых данных в XML
    for elem in root.iter():
        if 'LegalEntity' in elem.tag:
            # Обработка данных юридического лица
            person_type = 'ЮЛ'
            for elem_ul in elem.iter():
                if 'OrgFullNameUL' in elem_ul.tag:
                    org_name = elem_ul.text
                if 'ChiefUL' in elem_ul.tag:
                    # Формирование полного имени из элементов
                    full_name = ' '.join([e.text for e in elem_ul.iter() if e.text]).strip()
                if 'OrgCoordinatesUL' in elem_ul.tag:
                    # Извлечение контактной информации
                    for e in elem_ul.iter():
                        if 'Address' in e.tag:
                            address = e.text
                        if 'Phone' in e.tag:
                            phone = e.text
                        if 'Email' in e.tag:
                            email = e.text

        if 'IndividualEntrepreneur' in elem.tag:
            # Обработка данных индивидуального предпринимателя
            person_type = 'ИП'
            for elem_ip in elem.iter():
                if 'ChiefIP' in elem_ip.tag:
                    # Формирование полного имени из элементов
                    full_name = ' '.join([e.text for e in elem_ip.iter() if e.text]).strip()
                    org_name = 'ИП ' + full_name
                if 'OrgCoordinatesIP' in elem_ip.tag:
                    # Извлечение контактной информации
                    for e in elem_ip.iter():
                        if 'Address' in e.tag:
                            address = e.text
                        if 'Phone' in e.tag:
                            phone = e.text
                        if 'Email' in e.tag:
                            email = e.text

        if 'Person' in elem.tag:
            # Обработка данных физического лица
            person_type = 'ФЛ'
            for elem_fl in elem.iter():
                if 'LastnameFL' in elem_fl.tag:
                    last_name_fl = elem_fl.text
                if 'FirstnameFL' in elem_fl.tag:
                    first_name_fl = elem_fl.text
                if 'MiddlenameFL' in elem_fl.tag:
                    middle_name_fl = elem_fl.text
            # Формирование полного имени и названия организации для физического лица
            full_name = ' '.join([last_name_fl, first_name_fl, middle_name_fl])
            org_name = full_name

    # Возврат извлеченных данных
    return [application_number, person_type, full_name, org_name, address, phone, email]


def process_directory(root_dir):
    data = []
    # Рекурсивный обход директорий
    for dirpath, dirnames, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.endswith('.xml'):
                file_path = os.path.join(dirpath, filename)
                data.append(parse_xml(file_path))
    return data


# Обработка всех XML файлов в директории
result_data = process_directory(PATH_TO_DIR)

# Создание DataFrame из полученных данных
df = pd.DataFrame(result_data,
                  columns=['Номер заявки', 'Тип лица', 'ФИО', 'Наименование полное организации', 'Адрес',
                           'Номер телефона', 'E-Mail'])

# Сохранение результатов в файл
if OUT_FILE_EXTENTION == 'csv':
    df.to_csv(f'{OUT_FILE_NAME}.csv', index=False)
    print(f"Парсинг завершен. Результаты сохранены в файл '{OUT_FILE_NAME}.csv'")
else:
    df.to_excel(f'{OUT_FILE_NAME}.xlsx', index=False)
    print(f"Парсинг завершен. Результаты сохранены в файл '{OUT_FILE_NAME}.xlsx'")


