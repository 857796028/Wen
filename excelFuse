import xlrd
import xlwt
file1 = 'Book1.xls'
file2 = 'Book3.xls'
file_all = xlwt.Workbook()
fa_sheet1 = file_all.add_sheet('all')
wb1 = xlrd.open_workbook(file1)
wb2 = xlrd.open_workbook(file2)
sheet1 = wb1.sheet_by_index(0) #第一个表的第一个sheet
sheet2 = wb2.sheet_by_index(0) #第二个表的第一个sheet
sheet1_names = sheet1.row_values(0)  #属性
sheet2_names = sheet2.row_values(0)  #属性
sheet1_uindex = sheet1_names.index('name') #唯一属性索引
sheet2_uindex = sheet2_names.index('name') #唯一属性索引
sheet2_index1 = [] #sheet2独有的属性
all_names = sheet1_names #合并表的所有有属性
for i in sheet2_names:
    if i not in all_names:  
        all_names.append(i)
        sheet2_index1.append(sheet2_names.index(i))
#往sheet1后添加sheet2独有属性
for i in range(sheet1.nrows):
    d_r = sheet1.row_values(i)
    d_u = d_r[sheet1_uindex]
    d_c = sheet2.col_values(sheet2_uindex)
    d_x = d_c.index(d_u) if d_u in d_c else ' '#sheet2横坐标
    for j in range(sheet1.ncols + len(sheet2_index1)):
        if j < sheet1.ncols:
            f_value = sheet1.cell_value(i,j)
        else:
            d_y = sheet2_index1[j-sheet1.ncols] #sheet2纵坐标
            f_value = ' ' if d_x == ' ' else sheet2.cell_value(d_x,d_y)
        fa_sheet1.write(i,j,f_value)
r = sheet1.nrows
#w往sheet1下添加sheet2独有唯一属性
for m in range(1,sheet2.nrows):
    d_r = sheet2.row_values(m)
    d_u = d_r[sheet2_uindex]
    d_c = sheet1.col_values(sheet1_uindex)
    if d_u in d_c:
        continue
    for n in range(len(all_names)):
        d_s = all_names[n] #当前属性列     
        f_value = sheet2.cell_value(m,sheet2_names.index(d_s)) if d_s in sheet2_names else ' '
        fa_sheet1.write(r,n,f_value)
    r += 1
file_all.save('all1.xls')     
