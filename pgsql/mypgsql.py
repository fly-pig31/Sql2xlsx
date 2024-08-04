import psycopg2
import json
import os


class Conn:
    def __init__(self, db, usr, pwd, host, port, schema):
        self.db = db
        self.usr = usr
        self.pwd = pwd
        self.host = host
        self.port = port
        self.schema = schema

    def __del__(self):
        print(f"Object {self} is being destroyed.")


def schema_info(t_name: str, conn: Conn):
    con = psycopg2.connect(database=conn.db,
                           user=conn.usr,
                           password=conn.pwd,
                           host=conn.host,
                           port=conn.port)
    cur = con.cursor()
    schema = conn.schema

    cur.execute(f'''SET SCHEMA '{schema}' ''')
    cur.execute(f'''
        SELECT
            tab.table_comment AS "表注释",
            '{t_name}' AS "表名",
            cols.column_name AS "字段名",
            pgd.description AS "注释",
            cols.data_type AS "数据类型",
            cols.character_maximum_length AS "长度",
            cols.numeric_precision AS "小数位数",
            cols.is_nullable AS "是否允许为空",
            cols.column_default AS "默认值"
        FROM
            information_schema.columns AS cols
        LEFT JOIN
            pg_catalog.pg_description AS pgd ON cols.table_name::regclass = pgd.objoid
                AND pgd.objsubid = cols.ordinal_position
        LEFT JOIN
            (
                SELECT
                    pgc.oid AS table_oid,
                    pd.description AS table_comment
                FROM
                    pg_catalog.pg_class AS pgc
                LEFT JOIN
                    pg_catalog.pg_description AS pd ON pgc.oid = pd.objoid
                    AND pd.objsubid = 0
                WHERE
                    pgc.relname = '{t_name}'
            ) AS tab ON cols.table_name::regclass = tab.table_oid
        WHERE
            cols.table_name = '{t_name}' and cols.table_schema="{schema}"
        ''')
    return cur.fetchall()


def schema_info_s(tables: list, conn: Conn):
    data_name = conn.schema
    db_name = conn.db
    file_path = f"tmp_json/{db_name}_{data_name}_td.json"
    if not os.path.exists(file_path) or os.stat(file_path).st_size == 0:
        print("创建连接，连接数据库")
        con = psycopg2.connect(database=conn.db,
                               user=conn.usr,
                               password=conn.pwd,
                               host=conn.host,
                               port=conn.port)
        cur = con.cursor()
        schema = conn.schema
        result = []
        cur.execute(f'''SET SCHEMA '{schema}' ''')
        for t_name in tables:
            cur.execute(f'''
                SELECT
                    tab.table_comment AS "表注释",
                    '{t_name}' AS "表名",
                    cols.column_name AS "字段名",
                    pgd.description AS "注释",
                    cols.data_type AS "数据类型",
                    cols.character_maximum_length AS "长度",
                    cols.numeric_scale AS "小数位数",
                    cols.is_nullable AS "是否允许为空",
                    cols.column_default AS "默认值"
                FROM
                    information_schema.columns AS cols
                LEFT JOIN
                    pg_catalog.pg_description AS pgd ON cols.table_name::regclass = pgd.objoid
                        AND pgd.objsubid = cols.ordinal_position
                LEFT JOIN
                    (
                        SELECT
                            pgc.oid AS table_oid,
                            pd.description AS table_comment
                        FROM
                            pg_catalog.pg_class AS pgc
                        LEFT JOIN
                            pg_catalog.pg_description AS pd ON pgc.oid = pd.objoid
                            AND pd.objsubid = 0
                        WHERE
                            pgc.relname = '{t_name}'
                    ) AS tab ON cols.table_name::regclass = tab.table_oid
                WHERE
                    cols.table_name = '{t_name}'
                ''')
            print("执行查询")
            data = cur.fetchall()
            result.append(data)
        con.close()
        # print(result)
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False)

    print("正在使用缓存数据")
    with open(file_path, "r", encoding="utf-8") as file:
        json_data = json.load(file)
    return json_data
