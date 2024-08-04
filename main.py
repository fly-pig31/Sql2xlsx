from pgsql import mymysql
from pgsql.mymysql import MySQLConn
from xgui.write_result import write_to_word

"""
write_to_word
"""
input_tables = '''
db_user
'''
tables = input_tables.strip().split('\n')
conn = MySQLConn("db_name", "root", "root", "localhost", 3306)
result = mymysql.schema_info_s(tables, conn)
doc_path = "result/doc/tayipingyang.docx"
table_header_rows = ["字段名", "数据类型", "是否为空", "默认值", "注释"]
# xlsx_path = "/result/doc/result.xlsx"
write_to_word(result, doc_path, table_header_rows)
