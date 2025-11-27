from docx import Document


def replace_text(doc , find_text , replace_text):

    for para in doc.paragraphs:
        if find_text in para.text:
            para.text = para.text.replace(find_text, str(replace_text))


def replace_table(doc , find_text , replace_text):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if find_text in cell.text:
                    cell.text = cell.text.replace(find_text, str(replace_text))


