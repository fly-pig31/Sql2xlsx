import json

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL, WD_CELL_VERTICAL_ALIGNMENT
from pgsql.mypgsql import schema_info_s, Conn

conn = Conn("data", "postgres", "PPJbh9z7zrRrH9a4", "192.168.2.250", "5432", "sdxy")
input_string = '''
country_config
edu_app_sygk_byjyqk
edu_app_sygk_cjrh
edu_app_sygk_gjyx
edu_app_sygk_gl
edu_app_sygk_jyjg
edu_app_sygk_kyfw
edu_app_sygk_rcpy
edu_app_sygk_ssgk
edu_app_sygk_xxgk
edu_app_sygk_zykc
edu_app_bxtj_dbgczb
edu_app_bxtj_jcbxzb
edu_app_bxtj_jsqk
edu_app_bxtj_sbzb
edu_app_bxtj_tdjj
edu_app_bxtj_xsqk
edu_app_bxtj_xxzc
edu_app_xsgk_byqxqk
edu_app_xsgk_dekt
edu_app_xsgk_fwxszl
edu_app_xsgk_gfjygx
edu_app_xsgk_gl
edu_app_xsgk_jzcg
edu_app_xsgk_qjgk
edu_app_xsgk_qjgk_dqqj
edu_app_xsgk_stgm
edu_app_xsgk_xshjgk
edu_app_xsgk_xsjshj
edu_app_xsgk_zxs_mzfb
edu_app_xsgk_zxs_sydfb
edu_app_xsgk_zxs_zzmmfb
edu_app_xsgk_zxsgk_yx
edu_app_xsgk_zxsmzfb
edu_app_xsgk_zxsxbfb
edu_app_zhxq_bxtjgk
edu_app_zhxq_cwgk
edu_app_zhxq_djsz
edu_app_zhxq_djsz_djhdlx
edu_app_zhxq_djsz_szhdlx
edu_app_zhxq_gl
edu_app_zhxq_gl_zc
edu_app_zhxq_glfw
edu_app_zhxq_jkygk
edu_app_zhxq_jxgk
edu_app_zhxq_jyjx
edu_app_zhxq_kygk
edu_app_zhxq_kygk_kycg
edu_app_zhxq_rsgk
edu_app_zhxq_sp
edu_app_zhxq_spxk
edu_app_zhxq_spxk_copy1
edu_app_zhxq_sxjy
edu_app_zhxq_sxjy_fbqk
edu_app_zhxq_sxjy_hyfb
edu_app_zhxq_sysx
edu_app_zhxq_tsyy
edu_app_zhxq_xsgk
edu_app_zhxq_xsgl
edu_app_zhxq_xsgl_sthd
edu_app_rsgk_cbpx
edu_app_rsgk_fxjx
edu_app_rsgk_gccrc
edu_app_rsgk_gl
edu_app_rsgk_gzdl
edu_app_rsgk_hydsfx
edu_app_rsgk_jsgk
edu_app_rsgk_jzgrs_glry
edu_app_rsgk_jzgrs_zrjs
edu_app_rsgk_jzgzcfb
edu_app_rsgk_jzgzs
edu_app_rsgk_nlfb
edu_app_rsgk_ryfb
edu_app_rsgk_ssb
edu_app_rsgk_szdwcg
edu_app_rsgk_txjzgrs_glry
edu_app_rsgk_txjzgrs_zrjs
edu_app_rsgk_txjzgrs_zrjs_copy1
edu_app_rsgk_xlfb
edu_app_rsgk_xwjsrsbl
edu_app_rsgk_zcdjfb
edu_app_rsgk_zcfb_jf
edu_app_rsgk_zcfb_js
edu_app_rsgk_zrjs
edu_app_rsgk_zrjscjpx
edu_app_rsgk_zrjsqysj
edu_app_rsgk_ztqk
edu_app_rsgk_zzmmfb
edu_app_jxgk_gkszrt
edu_app_jxgk_gxyzs
edu_app_jxgk_jccg
edu_app_jxgk_jpzxkc
edu_app_jxgk_jxzl
edu_app_jxgk_kcjsqk
edu_app_jxgk_kcsz
edu_app_jxgk_sj_jxnldsqk
edu_app_jxgk_syfx
edu_app_jxgk_syjg
edu_app_jxgk_xjydqk
edu_app_jxgk_xnfz_jyjg
edu_app_jxgk_xsjshjqk
edu_app_jxgk_xsskqk
edu_app_jxgk_zsbksqk
edu_app_jxgk_zsgm
edu_app_jxgk_zsgmsj
edu_app_jxgk_zsgmzj
edu_app_jxgk_zygmpm
edu_app_jxgk_zyjs
edu_app_jxgk_zyjsfw
edu_app_jxgk_zyjsqk
edu_app_jxgk_zykcg
edu_app_jxgk_zzyfx
edu_app_kygk_bzyzcg
edu_app_kygk_gl
edu_app_kygk_hxkyxm
edu_app_kygk_hxkyxm_gl
edu_app_kygk_kycgj
edu_app_kygk_kycgj_nf
edu_app_kygk_kypt
edu_app_kygk_xqhz
edu_app_kygk_zscqsq
edu_app_kygk_zscqzr
edu_app_kygk_zxkyxm
edu_app_kygk_zxkyxm_gl
edu_app_zcgk_bmzcfb
edu_app_zcgk_dxyqsbslqk_dnxz
edu_app_zcgk_dxyqsbslqk_zl
edu_app_zcgk_gdzc_dnxz
edu_app_zcgk_gdzcfb
edu_app_zcgk_gdzcqk
edu_app_zcgk_gdzczj
edu_app_zcgk_gl
edu_app_zcgk_gxyjypm
edu_app_zcgk_jxkysb
edu_app_zcgk_nxzjxkysb
edu_app_zcgk_sjjypm
edu_app_zcgk_tszyfb
edu_app_zcgk_xnsxsqk
edu_app_zcgk_xsjzmjzb
edu_app_zcgk_yqsbsl
edu_app_zcgk_zcfbqk
edu_app_zcgk_zczl
edu_app_jlhz_dt
edu_app_jlhz_dt_lx
edu_app_jlhz_gjdqhzyx
edu_app_jlhz_gjhjdqd
edu_app_jlhz_gjjlhz
edu_app_jlhz_gl
edu_app_jlhz_xyjz
edu_app_jlhz_ygapxjd

'''

# 清除字符串中的空格和换行符，并按换行符分割成列表
tables = input_string.strip().split('\n')

# 使用集合去重，再转换为列表
tables = list(set(tables))


def write_to_word(result, filename):
    document = Document()

    # Define styles for table elements
    header_font_size = Pt(10)
    cell_font_size = Pt(10)

    # Loop through the result and create tables in the Word document
    for item in result:
        # print(f'item:{item}')
        table_name = item[0][0]
        table_comment = item[0][1]

        # Add an empty paragraph before each new table
        document.add_paragraph()

        # Create a new table for table name and comment
        table_header = document.add_table(rows=1, cols=1)
        table_header.width = Pt(500)
        table_header.style = 'Table Grid'

        # Add table name as a cell in the header table
        cell_name = table_header.cell(0, 0)
        cell_name.text = table_name if table_name else "N/A"  # Use "N/A" if table_name is None or empty
        cell_name.paragraphs[0].runs[0].font.size = header_font_size
        cell_name.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # Add table comment as a cell in the header table
        cell_comment = table_header.add_row().cells[0]
        cell_comment.text = table_comment.upper() if table_comment else "N/A"  # Use "N/A" if table_comment is None or empty
        cell_comment.paragraphs[0].runs[0].font.size = header_font_size
        cell_comment.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # Create a new table for the actual data
        table_data = document.add_table(rows=1, cols=7)
        table_data.width = Pt(500)
        table_data.style = 'Table Grid'

        # Set header row style
        header_row = table_data.rows[0]
        header_row.cells[0].text = "字段名"
        header_row.cells[1].text = "注释"
        header_row.cells[2].text = "数据类型"
        header_row.cells[3].text = "长度"
        header_row.cells[4].text = "小数位数"
        header_row.cells[5].text = "是否允许为空"
        header_row.cells[6].text = "默认值"

        for cell in header_row.cells:
            cell.paragraphs[0].runs[0].font.size = header_font_size
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # Set cell style for the rest of the rows
        for row in item[1:]:
            new_row = table_data.add_row().cells
            for i, cell_data in enumerate(row[2:]):
                new_row[i].text = str(cell_data).upper()
                new_row[i].paragraphs[0].runs[0].font.size = cell_font_size
                new_row[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    document.save(filename)


line = '''
---------------------------------------------------------------------------
'''
result = schema_info_s(tables, conn)
# sorted_result = sorted(result, key=lambda x: x[0][1].split("_")[2])
sorted_result = sorted(result, key=lambda x: x[0][1].split("_")[2] if len(x[0][1].split("_")) >= 3 else '')
# for row in sorted_result:
#     print(row)
print(f'选中的表:{tables}')
# for item in result:
#     for row in item:
#         print(f"row:{row}")
#     print(line)
print("Executing write function")
filename="闽江学院"
docx_name = f"{filename}.docx"
xlsx_name = f"{filename}.docx"
write_to_word(sorted_result, docx_name)
