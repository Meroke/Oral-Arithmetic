import os.path
import xlrd
import xlwt
import random
import time

sheet1_record = []
sheet2_record = []
sheet3_record = []
num = 1
# 行，列，最后列的行数
line1 =[0,0,0]
line2 = [0,0,0]
line3 = [0,0,0]

create_file_name = ''
file_check = True

# 旧算法核心，随即碰撞生成文件，耗时长，容易崩溃
# # 随机返回三下口算文件中第一列表的公式，乘法
# def get_sheet1(sheet,row_end,col_end):
#     global sheet1_record
#     col = random.randint(0, row_end)
#     row = random.randint(0, col_end)
#     # i = 0
#     while (row, col) in sheet1_record:
#         col = random.randint(0, row_end)
#         row = random.randint(0, col_end)
#     sheet1_record.append((row, col))
#     if sheet.cell(row, col).value:
#         return sheet.cell(row, col).value
#     else:
#         return None
#
#
# # 随机返回三下口算文件中第二列表的公式，除法
# def get_sheet2(sheet,row_end,col_end):
#     global sheet2_record
#     col = random.randint(0, row_end)
#     row = random.randint(0, col_end)
#     # i = 0
#     while (row, col) in sheet2_record:
#         # print('tip2::' + str(i))
#         # i += 1
#         col = random.randint(0, row_end)
#         row = random.randint(0, col_end)
#     sheet2_record.append((row, col))
#     if sheet.cell(row, col).value:
#         return sheet.cell(row, col).value
#     else:
#         return None
#
#
# # 随机返回三下口算文件中第三列表的公式，混合计算
# def get_sheet3(sheet,row_end,col_end):
#     global sheet3_record
#     col = random.randint(0, row_end)
#     row = random.randint(0, col_end)
#     # i = 0
#     while (row, col) in sheet3_record:
#         # print('tip3::' + str(i))
#         # i += 1
#         col = random.randint(0, row_end)
#         row = random.randint(0, col_end)
#     sheet3_record.append((row, col))
#     if sheet.cell(row, col).value:
#         return sheet.cell(row, col).value
#     else:
#         return None


# 新算法核心，采用序列排除，逐一挑选题目，生成文件，速度快，O(n)
def get_sheet_list(line):
    list = []
    for col_len in range(line[0]+1):

        for row_len in range(line[1] + 1):
        # for row_len in range(line[1]+1):
            if col_len == line[0] and row_len == line[2]:
                break
            list.append([row_len,col_len])
            # print(row_len,col_len)
    if list:
        return list
    else:
        return None

def get_sheet_way2(sheet,line):
    list_len = len(line) - 1
    # print('list_len: ' + str(list_len))
    magic_num = random.randint(0, list_len)
    order = line.pop(magic_num)
    # print(order)
    cell_value = sheet.cell(order[1], order[0]).value
    # print(cell_value)
    if cell_value:
        return cell_value
    else:
        return None


def get_line(sheet):
    row = sheet.row(0)
    col = sheet.col(0)
    # 一列的长度为行数，反之亦然
    col_end = sheet.col(len(row) -1) # 最后一列数据
    col_len = len(row) - 1
    row_len = len(col) - 1
    col_endlen = 0
    for i in range(len(col_end)):
        # print(i,col_len)
        # print(sheet.cell(i,col_len).value)
        if sheet.cell(i,col_len).ctype != 0:
            col_endlen += 1
    # col_endlen = len(col_end) - 1
    # print(row_len,col_len,col_endlen)

    return row_len,col_len,col_endlen-1

def get_Allsheets():
    book = xlrd.open_workbook('三下口算.xlsx')  # 读取原文件
    sheet1 = book.sheets()[0]  # 获取表一
    sheet2 = book.sheets()[1]  # 获取表二
    sheet3 = book.sheets()[2]  # 获取表三
    return sheet1, sheet2, sheet3

def check_fileOping(file_path):
    try:
        print(open(file_path,'w'))
        return True
    except Exception as e:
        if ("[Errno 13] Permission denied" in str(e)):
            print('file opening')
        return False



# 创建新的excel文件
def create_new_file(n1, n2, n3):

    sheet1,sheet2,sheet3 = get_Allsheets()

    global line1,line2,line3
    line1[0],line1[1],line1[2] = get_line(sheet1)
    line2[0],line2[1],line2[2] = get_line(sheet2)
    line3[0],line3[1],line3[2] = get_line(sheet3)
    # print(line1,line2,line3)
    # print(line1)
    line1_list = get_sheet_list(line1)
    line2_list = get_sheet_list(line2)
    line3_list = get_sheet_list(line3)
    # print(line1_list)

    R2 = n1 + n2  # 前两种题型的题数
    R3 = n1 + n2 + n3  # 总题数
    col_lines = R3 // 20 + 1

    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('test')

    font = xlwt.Font()
    font.height = 20 * 16

    style = xlwt.XFStyle()
    style.font = font

    col = 0
    row = 0
    temp = 0

    i = 0
    while i < n1:
    # for i in range(n1):
        col = int(i / 20)
        if col > temp:
            row = 0
        write_word = get_sheet_way2(sheet1,line1_list)
        if write_word is not None:
            worksheet.write(row, col, write_word, style)
            row += 1
            temp = col
            i += 1

    # i = n1
    while i < R2:
    # for i in range(n1, R2):
        col = int(i / 20)
        if col > temp:
            row = 0
        write_word = get_sheet_way2(sheet2,line2_list)
        if write_word is not None:
            worksheet.write(row, col, write_word, style)
            row += 1
            temp = col
            i += 1
    i = R2
    while i < R3:
    # for i in range(R2, R3):
        col = int(i / 20)
        if col > temp:
            row = 0
        write_word = get_sheet_way2(sheet3,line3_list)
        if write_word is not None:
            worksheet.write(row, col, write_word, style)
            row += 1
            temp = col
            i += 1

    # 设置列宽，按总题目数量
    for cols in range(0, col_lines):
        col = worksheet.col(cols)
        col.width = 256 * 30

    # 设置行高和字体大小，仅20行
    tall_style = xlwt.easyxf('font:height 650')
    for rows in range(0, 20):
        row = worksheet.row(rows)
        row.set_style(tall_style)

    # 设置文件名称 类型+序号
    file_name = "二和一"
    if n1 > 0 and n2 is 0 and n3 is 0:
        file_name = "乘法"
    elif n1 is 0 and n2 > 0 and n3 is 0:
        file_name = "除法"
    elif n1 is 0 and n2 is 0 and n3 > 0:
        file_name = "混合运算"
    elif n1 > 0 and n2 > 0 and n3 > 0:
        file_name = "三合一"
    global file_check
    file_check = True
    global num
    global create_file_name
    time_tuple = time.localtime(time.time())
    create_file_name = file_name + str(num) + '--' + \
                       str(time_tuple.tm_year) + '-' + str(time_tuple.tm_mon) + '-' + str(time_tuple.tm_mday) +'_' +  \
                       str(time_tuple.tm_hour) + '-' + str(time_tuple.tm_min) + '-' + str(time_tuple.tm_sec) + '.xls'
    if not os.path.exists(create_file_name):
        workbook.save(create_file_name)
        num += 1
    else:
        file_check = check_fileOping(create_file_name)

    # 清除残留记录
    sheet1_record.clear()
    sheet2_record.clear()
    sheet3_record.clear()


# 单文件测试，连续生成多个三合一文件
if __name__ == '__main__':
    for j in range(1):
        create_new_file(200, 140, 60)
