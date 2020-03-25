#练习代码
import xlrd
import xlwt
minsup=0.2
minconfi=0.5
data=xlrd.open_workbook(r'C:\Users\18765511549\Desktop\文件\练习\CatalogCrossSell.xls')
table=data.sheets()[0]
rows_num=table.nrows
cols_num=table.ncols
