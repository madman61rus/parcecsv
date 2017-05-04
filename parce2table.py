from docx import Document


def get_all_tables(filename):
    document = Document(filename)
    tables = []

    for table in document.tables:
        tables.append(table)

    return tables


def get_tables_by_number(tables=None, numbers=None):

    if tables is None and numbers is None:
        return []

    if numbers is None:
        return tables

    return [tables[number] for number in numbers]


def extract_data_from_table(table):
    data = []
    keys = None

    for i,row in enumerate(table[0].rows):
        text = (cell.text for cell in row.cells)

        if i == 0:
            keys = tuple(text)
            continue

        row_data = dict(zip(keys, text))
        data.append(row_data)

    return data


if __name__ == '__main__':
    file = 'ege2015.docx'
    tables_from_file = get_all_tables(file)
    print(get_tables_by_number(tables_from_file, [67]))
    print(extract_data_from_table(get_tables_by_number(tables_from_file, [67])))