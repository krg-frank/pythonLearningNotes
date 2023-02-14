import xlrd
import xlwt
from pathlib import Path, PurePath
# 导入excel和文件操作库

#查询是否匹配函数
def matchValue(data):
  # 遍历的范围 需确定
  matchedNum=[7,23,1,39,18,33,-1,16,6,32,11,26,8,22,16,5,10,24,38,17,32,2,21,6,40,15,19,34]
  #my_list=list(matchedNum)
  for i in range(len(matchedNum)-1):
    value=abs(data-matchedNum[i])
    if(value<=1):
      # print("matchedNum=",matchedNum[i],"data=",data)
      # print("my_list.len",len(matchedNum))
      return True
    
# #基于已知的匹配数做差生成二级匹配数
# src1_name=r'E:\desktop\MSDataAna\matchedNum.xls'
# dst1_name=r'E:\desktop\MSDataAna\matchedNum.xls'

# data = xlrd.open_workbook(src1_name)
# table = data.sheets()[0]

# row=1
# col = 0
# workbook = xlwt.Workbook(encoding='utf-8')
# xlsheet = workbook.add_sheet("Result")
# #准备写入文件的表头
# table_header = ['value','data1', 'data2']
# for cell_header in table_header:
#   xlsheet.write(row, col, cell_header)
#   col += 1 
# ncols = table.ncols
# for i in  range(0,ncols-1):
#   data1 = table.cell_value(rowx=0, colx=i)
#   for j in  range(i+1,ncols):
#     data2 = table.cell_value(rowx=0, colx=j)
#     value=abs(data2-data1)
#     #新增一个行
#     row += 1
#     col=0
#     xlsheet.write(row, col, value)
#     xlsheet.write(row, col+1, data1)
#     xlsheet.write(row, col+2, data2)

# # 保存最终结果
# workbook.save(dst1_name)





# #基于已知的匹配数做差生成二级匹配数
# src1_name=r'E:\desktop\MSDataAna\resultThd.xls'
# dst1_name=r'E:\desktop\MSDataAna\resultFor.xls'

# data = xlrd.open_workbook(src1_name)
# table = data.sheets()[0]

# row=0
# col = 0
# workbook = xlwt.Workbook(encoding='utf-8')
# xlsheet = workbook.add_sheet("Result")
# #准备写入文件的表头
# table_header = [ 'data1','data2','value']
# for cell_header in table_header:
#   xlsheet.write(row, col, cell_header)
#   col += 1 
# nrows = table.nrows
# for i in  range(1,nrows-1):
#   data1 = table.cell_value(rowx=i, colx=7)
#   for j in  range(i+1,nrows):
#     data2 = table.cell_value(rowx=j, colx=7)
#     value=data2-data1
#     #新增一个行
#     row += 1
#     col=0
#     xlsheet.write(row, col, data1)
#     xlsheet.write(row, col+1, data2)
#     xlsheet.write(row, col+2, -value)

# # 保存最终结果
# workbook.save(dst1_name)




# src_name=r'E:\desktop\MSDataAna\source.xlsx'
# dst_name=r'E:\desktop\MSDataAna\result.xls'

# data = xlrd.open_workbook(src_name)
# table = data.sheets()[0]

# row=0
# col = 0
# workbook = xlwt.Workbook(encoding='utf-8')
# xlsheet = workbook.add_sheet("matchResult")
# # 准备写入文件的表头
# table_header = ['data1', 'data2', 'value']
# for cell_header in table_header:
#   xlsheet.write(row, col, cell_header)
#   col += 1 
# ncols = table.ncols
# for i in  range(0,ncols-1):
#   data1 = table.cell_value(rowx=0, colx=i)
#   for j in  range(i+1,ncols):
#     data2 = table.cell_value(rowx=0, colx=j)
#     value=abs(data2-data1)
#     if(matchValue(value)):
#       #新增一个行
#       row += 1
#       col=0
#       xlsheet.write(row, col, data1)
#       xlsheet.write(row, col+1, data2)
#       xlsheet.write(row, col+2, data2-data1)

# # 保存最终结果
# workbook.save(dst_name)




# src3_name=r'E:\desktop\MSDataAna\result.xls'
# dst3_name=r'E:\desktop\MSDataAna\resultThd.xls'

# data = xlrd.open_workbook(src3_name)
# table = data.sheets()[0]

# row=0
# col = 0
# workbook = xlwt.Workbook(encoding='utf-8')
# xlsheet = workbook.add_sheet("Result")

# # 准备写入文件的表头
# table_header = ['value','c', 'd','a','b','M']
# for cell_header in table_header:
#   xlsheet.write(row, col, cell_header)
#   col += 1 

# nrows = table.nrows
# for i in  range(1,nrows):
#   data1 = table.cell_value(rowx=i, colx=3)
#   for j in  range(1,nrows):
#     data2 = table.cell_value(rowx=j, colx=4)
#     if(data1==data2):
#       rowValue=table.row_values(j)
#       a=rowValue[5]
#       b=rowValue[6]
#       rowValue=table.row_values(i)
#       c=rowValue[0]
#       d=rowValue[1]
#       # print("a,b,c,d=",a,b,c,d)
#       if(c-a==d-b):
#         row += 1
#         col=0
#         xlsheet.write(row, col, data1)
#         xlsheet.write(row, col+1, c)
#         xlsheet.write(row, col+2, d)
#         xlsheet.write(row, col+3, a)
#         xlsheet.write(row, col+4, b)
#         xlsheet.write(row, col+5, c-a)
#       if(d-a==c-b):
#         row += 1
#         col=0
#         xlsheet.write(row, col, data1)
#         xlsheet.write(row, col+1, c)
#         xlsheet.write(row, col+2, d)
#         xlsheet.write(row, col+3, a)
#         xlsheet.write(row, col+4, b)
#         xlsheet.write(row, col+5, d-a)

# # 保存最终结果
# workbook.save(dst3_name)




# src3_name=r'E:\desktop\MSDataAna\resultFor.xls'
# dst3_name=r'E:\desktop\MSDataAna\resultFir.xls'

# data = xlrd.open_workbook(src3_name)
# table = data.sheets()[0]

# row=0
# col = 0
# workbook = xlwt.Workbook(encoding='utf-8')
# xlsheet = workbook.add_sheet("Result")

# # 准备写入文件的表头
# table_header = ['value','data1', 'data2']
# for cell_header in table_header:
#   xlsheet.write(row, col, cell_header)
#   col += 1 

# nrows = table.nrows
# for i in  range(1,nrows):
#   data1 = table.cell_value(rowx=i, colx=2)
#   for j in  range(1,nrows):
#     data2 = table.cell_value(rowx=j, colx=3)
#     if(data1==data2):
#       rowValue=table.row_values(i)
#       c=rowValue[0]
#       d=rowValue[1]
#       row += 1
#       col=0
#       xlsheet.write(row, col, data1)
#       xlsheet.write(row, col+1, c)
#       xlsheet.write(row, col+2, d)

# # 保存最终结果
# workbook.save(dst3_name)


# src3_name=r'E:\desktop\MSDataAna\resultFir.xls'
# dst3_name=r'E:\desktop\MSDataAna\resultSix.xls'

# data = xlrd.open_workbook(src3_name)
# table = data.sheets()[0]

# row=0
# col = 0
# workbook = xlwt.Workbook(encoding='utf-8')
# xlsheet = workbook.add_sheet("Result", cell_overwrite_ok=True)
# # worksheet = workbook.add_sheet("Sheet 1")

# # 准备写入文件的表头
# table_header = ['value','data1', 'data2']
# for cell_header in table_header:
#   xlsheet.write(row, col, cell_header)
#   col += 1 

# #############################################################
# ## 行之间进行查重 未完成
# #############################################################
# nrows = table.nrows
# for i in  range(1,nrows-1):
#   data1 = table.cell_value(rowx=i, colx=0)
#   data2 = table.cell_value(rowx=i, colx=1)
#   data3 = table.cell_value(rowx=i, colx=2)
#   for j in  range(i+1,nrows):
#     data4 = table.cell_value(rowx=j, colx=0)
#     data5 = table.cell_value(rowx=j, colx=1)
#     data6 = table.cell_value(rowx=j, colx=2)
#     if(data1==data4 and data2==data5 and data3==data6 and data4 >=0):
#       # rowValue=table.row_values(i)
#       # c=rowValue[0]
#       # d=rowValue[1]
#       # row += 1
#       # col=0
#       xlsheet.write(j, 0, -1)
#       xlsheet.write(j, 1, -1)
#       xlsheet.write(j, 2, -1)

# # 保存最终结果
# workbook.save(dst3_name)





src5_name=r'E:\desktop\MSDataAna\resultFir.xls'
src6_name=r'E:\desktop\MSDataAna\resultSix.xls'
dst_name=r'E:\desktop\MSDataAna\resultSev.xls'

data5 = xlrd.open_workbook(src5_name)
table5 = data5.sheets()[0]
data6 = xlrd.open_workbook(src6_name)
table6 = data6.sheets()[0]

row=0
col = 0
workbook = xlwt.Workbook(encoding='utf-8')
xlsheet = workbook.add_sheet("Result", cell_overwrite_ok=True)
# worksheet = workbook.add_sheet("Sheet 1")

# 准备写入文件的表头
table_header = ['value','data1', 'data2']
for cell_header in table_header:
  xlsheet.write(row, col, cell_header)
  col += 1 

nrows = table6.nrows
for i in  range(1,nrows-1):
  data1 = table6.cell_value(rowx=i, colx=0)
  if(data1==""):
    data2 = table5.cell_value(rowx=i, colx=0)
    data3 = table5.cell_value(rowx=i, colx=1)
    data4 = table5.cell_value(rowx=i, colx=2)
    xlsheet.write(i, 0, data2)
    xlsheet.write(i, 1, data3)
    xlsheet.write(i, 2, data4)
  else:
    data2 = table6.cell_value(rowx=i, colx=0)
    data3 = table6.cell_value(rowx=i, colx=1)
    data4 = table6.cell_value(rowx=i, colx=2)
    xlsheet.write(i, 0, data2)
    xlsheet.write(i, 1, data3)
    xlsheet.write(i, 2, data4)
# 保存最终结果
workbook.save(dst_name)