import os
import comtypes.client


def doc_to_docx(doc_path, docx_path):
    # 启动Word应用程序
    word = comtypes.client.CreateObject('Word.Application')

    # 打开.doc文件
    doc = word.Documents.Open(doc_path)

    # 保存为.docx文件
    doc.SaveAs(docx_path, FileFormat=16)

    # 关闭.doc文件
    doc.Close()

    # 退出Word应用程序
    word.Quit()


# 指定包含.doc文件的文件夹路径
doc_folder = r'D:\Study\Python\PycharmProjects\Sql2xlsx\file_handle\resources\111'

# 指定保存.docx文件的文件夹路径
docx_folder = r'D:\Study\Python\PycharmProjects\Sql2xlsx\file_handle\resources\111'

# 遍历文件夹中的所有.doc文件并转换为.docx
for filename in os.listdir(doc_folder):
    if filename.endswith(".doc"):
        doc_path = os.path.join(doc_folder, filename)
        docx_path = os.path.join(docx_folder, filename.replace(".doc", ".docx"))
        doc_to_docx(doc_path, docx_path)
        print(filename + " 转换完成！")

print("批量转换完成！")