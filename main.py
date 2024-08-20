from lxml import etree
import openpyxl


def parse_xsd_to_excel(xsd_file, output_file):
    # Парсинг XML схемы
    tree = etree.parse(xsd_file)
    root = tree.getroot()

    # Создаем Excel файл и добавляем заголовки
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "XSD Elements"
    ws.append(["Номер по порядку", "Название элемента", "Описание элемента", "Признак обязательности",
               "Повторяемость", "XPath", "Пояснение MDR", "Пояснение НРД"])

    # Счетчик для строки
    row_number = 1

    # Рекурсивная функция для обработки элементов
    def process_element(element, xpath=""):
        nonlocal row_number

        # Получение информации об элементе
        name = element.attrib.get('name')
        print(f"name:  {name} ")
        description = ""
        required = "Н"
        repeatable = "1"
        mdr = ""
        nsdr = ""

        xpath += f"/{name}" if name else ""

        # Проверка обязательности и повторяемости
        if element.tag.endswith('element'):
            if element.attrib.get('minOccurs', '1') == '0':
                required = "Н"
            else:
                required = "О"

            repeatable = element.attrib.get('maxOccurs', '1')

        # Получение аннотаций
        for annotation in element.findall(".//{http://www.w3.org/2001/XMLSchema}annotation"):
            for documentation in annotation.findall(".//{http://www.w3.org/2001/XMLSchema}documentation"):
                lang = documentation.attrib.get("{http://www.w3.org/XML/1998/namespace}lang")
                if lang == "RusEng":
                    mdr = documentation.text.strip() if documentation.text else ""
                elif lang == "Rus":
                    nsdr = documentation.text.strip() if documentation.text else ""

        # Запись строки в Excel
        if name:
            row_number += 1
            ws.append([row_number, name, description, required, repeatable, xpath, mdr, nsdr])

        # Рекурсивный обход дочерних элементов
        for child in element:
            process_element(child, xpath)

    # Начинаем обработку с корневого элемента
    for element in root.findall(".//{http://www.w3.org/2001/XMLSchema}element"):
        process_element(element)

    # Сохраняем Excel файл
    wb.save(output_file)
    print(f"Excel файл '{output_file}' успешно создан.")


# Использование функции
xsd_file = 'camt.053.001.06.xsd'
output_file = 'nalogadmin.xlsx'
parse_xsd_to_excel(xsd_file, output_file)
