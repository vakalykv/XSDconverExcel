import xml.etree.ElementTree as ET
import pandas as pd

# Parse the XSD file
tree = ET.parse(xsd_file_path)
root = tree.getroot()

# XML Schema namespace
xs_namespace = {'xs': 'http://www.w3.org/2001/XMLSchema'}

# Initialize the DataFrame to hold the extracted data
data = []
counter = 0


# Function to extract element details recursively
def process_element(element, parent_xpath):
    global counter

    # Extract element name and xpath
    name = element.get('name')
    xpath = f"{parent_xpath}/{name}" if parent_xpath else name

    # Extract type and min/max occurs
    type_name = element.get('type')
    min_occurs = element.get('minOccurs', '1')
    max_occurs = element.get('maxOccurs', '1')

    # Determine mandatory and repeatable flags
    mandatory = 'O' if min_occurs == '1' else 'N'
    repeatable = 'Yes' if max_occurs == 'unbounded' else 'No'

    # Initialize annotations
    mdr_annotation = ""
    nsdr_annotation = ""

    # Extract annotations, if any
    for annotation in element.findall('xs:annotation/xs:documentation', xs_namespace):
        if annotation.get('source') == 'MDR':
            mdr_annotation = annotation.text.strip()
        elif annotation.get('source') == 'NSDR':
            nsdr_annotation = annotation.text.strip()

    # Extract the description (if available) and type information
    description = element.find('xs:annotation/xs:documentation[@xml:lang="RusEng"]', xs_namespace)
    description_text = description.text.strip() if description is not None else ""

    # Handle type definition if provided in-line (simpleType/complexType)
    if type_name is None:
        simple_type = element.find('xs:simpleType', xs_namespace)
        complex_type = element.find('xs:complexType', xs_namespace)

        # Process simpleType enumerations if available
        if simple_type is not None:
            type_name = 'simpleType'
            enumerations = simple_type.findall('xs:restriction/xs:enumeration', xs_namespace)
            enumeration_values = [enum.get('value') for enum in enumerations]
            if enumeration_values:
                type_name += f" (Enumerations: {', '.join(enumeration_values)})"
        elif complex_type is not None:
            type_name = 'complexType'

    # Append extracted information to data
    counter += 1
    data.append(
        [counter, name, description_text, mandatory, repeatable, xpath, mdr_annotation, nsdr_annotation, type_name])

    # Process child elements if they exist (for complexType definitions)
    for child in element.findall('xs:complexType/xs:sequence/xs:element', xs_namespace):
        process_element(child, xpath)


# Start processing from the root elements
for elem in root.findall('xs:element', xs_namespace):
    process_element(elem, '')

# Convert the data to a DataFrame
df = pd.DataFrame(data, columns=['Номер по порядку', 'Название элемента', 'Описание элемента', 'Признак обязательности',
                                 'Повторяемость', 'XPath', 'Пояснение MDR', 'Пояснение НРД', 'Тип данных'])

# Save to Excel file
output_excel_path = 'xsd_to_excel_output.xlsx'
df.to_excel(output_excel_path, index=False)

# Show the first few rows of the resulting dataframe
df.head()

# Основной код
if __name__ == "__main__":
    # Путь к XML схеме
    xml_schema_file = "camt.053.001.06.xsd"

    # Парсим XML схему
    elements = parse_xml_schema(xml_schema_file)

    # Создаем Excel файл
    output_excel_file = "nalogadmin.xlsx"
    create_excel(elements, output_excel_file)

    print(f"Файл {output_excel_file} успешно создан.")
