
from docx import Document

def read_docx_tables(file_path):
    doc = Document(file_path)
    tables = []
    for table in doc.tables:
        for row in table.rows:
            row_text = [cell.text for cell in row.cells]
            tables.append(row_text)
    return tables

# 指定要读取的 Word 文档路径
docx_file = 'resources/scanner/10.20.4.1.docx'

# 读取 Word 文档中的表格内容
doc_tables = read_docx_tables(docx_file)
for table in doc_tables:
    print(table)