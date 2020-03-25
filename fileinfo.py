#练习代码
import xlrd
import xlwt
minsup=0.2
minconfi=0.5
data=xlrd.open_workbook(r'C:\Users\18765511549\Desktop\文件\练习\CatalogCrossSell.xls')
table=data.sheets()[0]
rows_num=table.nrows
cols_num=table.ncols
a=[]
ann=[]
an={}
#a集合将9种商品按顺序排列
for i in range(cols_num):
    a.append(table.row(0)[i].value)
#ai为一一映射的字典，即商品组合——支持度   aii为满足最小支持度的商品组合
#开始
a1={}
a11=[]
for i in range(0,cols_num):
    b=table.col_values(i)
    if (b.count(1)/(rows_num-1)>=minsup):
        a11.append(table.row(0)[i].value)
        a1[table.row(0)[i].value]=round(b.count(1)/(rows_num-1),4)
#上面一商品组合支持度符合条件
ann=ann+a11
for z in range(2,cols_num+1):
    if z==2:
        aii=a11
        ai=a1
    if len(aii)>=2:
        for i in range(len(aii)):
            for j in range(i+1,len(aii)):
                c2=[]
                for v in aii[i]:
                    if v in aii[j]:
                        c2.append(v)
                c3=[]
                if len(c2)==len(aii)+1:
                    for p in range(len(c2)):
                        c3.append(a.index(c2[p]))
                    num=0
                    c1=[]
                    print(c3)
                    for t in range(1,rows_num):
                        num1=0
                        for e in range(len(c2)):
                            if table.row(t)[c3[e]].value==1:
                                num1+=1
                        if num1/len[c2]==1:
                            num+=1
                    if num/(rows_num-1)>=0.2:
                        aii=[]
                        ai={}
                        for k in range(len(c2)):
                            c1.append(table.row(0)[c3[k]].value)
                        if c1 not in aii:
                            aii.append(c1)
                            ai[tuple(c1)]=round(num/(rows_num-1),4)
        ann=ann+aii
        an.update(ai)
print(an)
print(ann)
