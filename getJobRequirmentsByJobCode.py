import xlrd
from xlutils.copy import copy

padding_dict = {}
def read_excel_xls(path):
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    #定位行列
    for i in range(0, worksheet.nrows):
        for j in range(0, worksheet.ncols):
            if str(worksheet.cell_value(i,j)).strip() == "岗位代码":
                gwdm_col = j
                gwdm_row = i
            if str(worksheet.cell_value(i,j)).strip() == "年龄":
                nl_col = j
                nl_row = i
            if str(worksheet.cell_value(i,j)).strip() == "学历":
                xl_col = j
                xl_row = i
            if str(worksheet.cell_value(i,j)).strip() == "学位":
                xw_col = j
                xw_row = i
            if str(worksheet.cell_value(i,j)).strip() == "职称":
                zc_col = j
                zc_row = i
            if str(worksheet.cell_value(i,j)).strip() == "专业及代码":
                zydm_col = j
                zydm_row = i

    for i in range(nl_row+1, worksheet.nrows):
        gwdm = worksheet.cell_value(i, gwdm_col)
        nl = worksheet.cell_value(i, nl_col)
        xl = worksheet.cell_value(i, xl_col)
        xw = worksheet.cell_value(i,xw_col)
        zc = worksheet.cell_value(i,zc_col)
        zydm = worksheet.cell_value(i,zydm_col)

        padding_dict[gwdm] = nl, xl, xw, zc, zydm




def write_excel_xls_append(path):
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格

    # 定位行列
    for i in range(0, worksheet.nrows):
        for j in range(0, worksheet.ncols):
            if str(worksheet.cell_value(i, j)).strip() == "岗位代码":
                gwdm_col = j
                gwdm_row = i
            if str(worksheet.cell_value(i, j)).strip() == "年龄":
                nl_col = j
                nl_row = i
            if str(worksheet.cell_value(i, j)).strip() == "学历":
                xl_col = j
                xl_row = i
            if str(worksheet.cell_value(i, j)).strip() == "学位":
                xw_col = j
                xw_row = i
            if str(worksheet.cell_value(i, j)).strip() == "职称":
                zc_col = j
                zc_row = i
            if str(worksheet.cell_value(i, j)).strip() == "专业及代码":
                zydm_col = j
                zydm_row = i
    # 获取表格中已存在的nl,xl,xw,zc,zydm数据的行数
    rows_old =worksheet.nrows
    for i in range(nl_row + 1, worksheet.nrows):
        if worksheet.cell_value(i,nl_col) == '':
            rows_old = i
            break
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(rows_old, worksheet.nrows):
        new_worksheet.write(i, nl_col, padding_dict[str(worksheet.cell_value(i, gwdm_col)).strip()][0])
        new_worksheet.write(i, xl_col, padding_dict[str(worksheet.cell_value(i, gwdm_col)).strip()][1])
        new_worksheet.write(i, xw_col, padding_dict[str(worksheet.cell_value(i, gwdm_col)).strip()][2])
        new_worksheet.write(i, zc_col, padding_dict[str(worksheet.cell_value(i, gwdm_col)).strip()][3])
        new_worksheet.write(i, zydm_col, padding_dict[str(worksheet.cell_value(i, gwdm_col)).strip()][4])
    new_workbook.save(path)  # 保存工作簿
    print("xls格式表格【追加】写入数据成功！")


# read_excel_xls("C:/Users/Administrator/Desktop/f.xls")
# print(padding_dict)
# write_excel_xls_append("C:/Users/Administrator/Desktop/d.xls")

while True:
    read_path = input("岗位表路径:")
    write_path = input("人员表路径:")
    read_excel_xls(read_path)
    write_excel_xls_append(write_path)
