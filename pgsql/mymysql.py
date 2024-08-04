import pymysql


class MySQLConn:
    def __init__(self, db, usr, pwd, host, port):
        self.db = db
        self.usr = usr
        self.pwd = pwd
        self.host = host
        self.port = port

    def __del__(self):
        print(f"Object {self} is being destroyed.")


def schema_info_s(tables: list, conn: MySQLConn):
    with pymysql.connect(host=conn.host, user=conn.usr, password=conn.pwd, database=conn.db, port=conn.port) as con:
        cursor = con.cursor()
        result = []
        for t_name in tables:
            cursor.execute(f'''
            SELECT
            c.TABLE_NAME AS '表名',
            t.TABLE_COMMENT AS '表备注',
            c.COLUMN_NAME AS '字段名',
            c.COLUMN_TYPE AS '数据类型',
            c.IS_NULLABLE AS '是否允许为空',
            c.COLUMN_DEFAULT AS '默认值',
            c.COLUMN_COMMENT AS '注释' 
            FROM
            INFORMATION_SCHEMA.COLUMNS c
            INNER JOIN INFORMATION_SCHEMA.TABLES t ON c.TABLE_SCHEMA = t.TABLE_SCHEMA AND c.TABLE_NAME = t.TABLE_NAME
            WHERE
            c.TABLE_SCHEMA = '{conn.db}' -- 替换为实际的数据库名称
            AND c.TABLE_NAME = '{t_name}';
            ''')
            print("执行查询")
            data = cursor.fetchall()
            result.append(data)
    return result


# 使用示例
conn = MySQLConn("java_shop", "root", "root", "localhost", 3306)
result = schema_info_s(["db_user"], conn)
print(result)
