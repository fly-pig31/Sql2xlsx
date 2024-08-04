def merge_every_two_lines(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    merged_lines = []
    previous_line = ''
    for line in lines:
        if line.strip():  # 如果不是空行
            if previous_line.strip():  # 如果前一行也不是空行
                merged_lines.append(previous_line.strip() + ' ' + line.strip())  # 将前一行和当前行合并为一行
            else:
                merged_lines.append(line.strip())  # 否则只添加当前行
        previous_line = line

    with open(output_file, 'w', encoding='utf-8') as file:
        for line in merged_lines:
            file.write(line + '\n')


# 使用示例
input_file = r'C:\Users\FrederickBarbarrossa\Desktop\秦妇吟.txt'
output_file = 'output.txt'
merge_every_two_lines(input_file, output_file)
