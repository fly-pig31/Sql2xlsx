from pgsql.mypgsql import schema_info_s, Conn
# from pgsql.mymysql import schema_info_s, MySQLConn
from sql2docu import write_to_excel
from sql2word import write_to_word

conn = Conn("data", "postgres", "pwd", "192.168.1.1", "5432", "xhz")
# MySQLConn("data", "root", "pwd", "192.168.1.1", "3306", "xhz")
# 输入表名
input_string = '''

'''

# 清除字符串中的空格和换行符，并按换行符分割成列表
tables = input_string.strip().split('\n')

# 使用集合去重，再转换为列表
tables = list(set(tables))
result = schema_info_s(tables, conn)
for item in result:
    for row in item:
        print(f"row:{row}")
    # print(line)
print("执行写入函数")
filename = "excel_name"
docx_name = f"{filename}.docx"
xlsx_name = f"{filename}.xlsx"
# 导出到excel
write_to_excel(result, xlsx_name)
# 导出到word
# write_to_word(result, docx_name)
