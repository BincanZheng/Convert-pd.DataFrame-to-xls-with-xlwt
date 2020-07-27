#!/usr/bin/env python
# coding: utf-8

# # Imports

# In[1]:


import pandas as pd
import numpy as np
import xlwt
import xlrd


# # Data Read

# In[2]:


monthly_pandas_data = pd.read_csv(r'.\Sample Data.csv',encoding='gbk',index_col=0,mangle_dupe_cols=2,header=[0,1])


# # Help Functions

# ## Cell Formater

# In[3]:


def _get_style(borders_major='tblr',width_major=1,width_minor=1,font_size=10):
    '''
    borders_major：选择主要边框 默认为4边全选 边框粗细为 width_major 
    width_major：主要边框的粗细 默认为1
    wdith_minor：次要边框的粗细 默认为1
    font_size：字体大小 默认为10
    '''
    style = xlwt.XFStyle()       # Create Style
    font = xlwt.Font()           # Create Font
    borders = xlwt.Borders()     # Create Borders
    alignment = xlwt.Alignment() # Create Alignment
    font.name = '宋体'           # 设置字体为 宋体
    font.height = font_size*20   # 设置字体大小为 10（10*20）
    alignment.horz = xlwt.Alignment.HORZ_CENTER 
    # 可以选择: HORZ_GENERAL,HORZ_LEFT,HORZ_CENTER,HORZ_RIGHT,HORZ_FILLED,
    #          HORZ_JUSTIFIED,HORZ_CENTER_ACROSS_SEL,HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER 
    # 可以选择: VERT_TOP,VERT_CENTER,VERT_BOTTOM,VERT_JUSTIFIED,VERT_DISTRIBUTED
    alignment.wrap = 1           # 自动换行
    # 设置边框宽度
    borders.left = width_major if 'l' in borders_major else width_minor
    borders.right = width_major if 'r' in borders_major else width_minor
    borders.top = width_major if 't' in borders_major else width_minor
    borders.bottom = width_major if 'b' in borders_major else width_minor
    # 向style输入格式
    style.font = font
    style.alignment = alignment
    style.borders = borders
    return style


# In[4]:


def _get_style_2(l=1,r=1,b=1,t=1,font_size=20):
    '''
    l：l for left  ,左边框的粗细 默认为1
    r：r for right ,右边框的粗细 默认为1
    b：b for bottom,下边框的粗细 默认为1
    t：t for top   ,上边框的粗细 默认为1
    font_size：字体大小 默认为20
    '''
    style = xlwt.XFStyle()       # Create Style
    font = xlwt.Font()           # Create Font
    borders = xlwt.Borders()     # Create Borders
    alignment = xlwt.Alignment() # Create Alignment
    font.name = '宋体'           # 设置字体为 宋体
    font.height = font_size*20   # 设置字体大小为 20（20*20）
    alignment.horz = xlwt.Alignment.HORZ_CENTER 
    # 可以选择: HORZ_GENERAL,HORZ_LEFT,HORZ_CENTER,HORZ_RIGHT,HORZ_FILLED,
    #          HORZ_JUSTIFIED,HORZ_CENTER_ACROSS_SEL,HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER 
    # 可以选择: VERT_TOP,VERT_CENTER,VERT_BOTTOM,VERT_JUSTIFIED,VERT_DISTRIBUTED
    # 设置边框宽度
    borders.left = l
    borders.right = r
    borders.top = t
    borders.bottom = b
    # 向style输入格式
    style.font = font
    style.alignment = alignment
    style.borders = borders
    return style


# ## XLWT write valid value

# In[5]:


def _Int_Print(value):
    '''
    is number  -> int(value)
    is nan     -> ''
    not number -> value
    '''
    try:
        if str(value).isnumeric():
            return int(value)
        elif np.isnan(value):
            return ''
        else:
            return int(value)
    except:
        return value


# # Excel Write

# ## Cell Formats Preparation

# In[6]:


# 表头
font_style_lrtb_0   = _get_style('lrtb',0)
font_style_lrtb     = _get_style('lrtb',2)
font_style_lrt      = _get_style('lrt',2)
font_style_lrb      = _get_style('lrb',2)
font_style_rtb      = _get_style('rtb',2)
font_style_b        = _get_style('b',2)
font_style_lb        = _get_style('lb',2)
font_style_rb        = _get_style('rb',2)

# 表内容
font_style_lr       = _get_style('lr',2)
# font_style_lrb      = _get_style('lrb',2)
font_style_r        = _get_style('r',2)
# font_style_rb       = _get_style('rb',2)
font_style_normal   = _get_style()
# font_style_b        = _get_style('b',2)


# ## Normal Settings

# In[7]:


normal_data = '空  气  温  度  （0.1℃）'   # 一级 column name
avg_data = '平均'                           # 一级&二级 column name
max_data = '最高'                           # 一级&二级 column name
min_data = '最低'                           # 一级&二级 column name
year = str(2019)                            # 使用参数
month = '{:02}'.format(2)                   # 使用参数

# 展示的数据名为monthly_excel(pandas的DataFrame形式)
# 把index放入column（写入表格时使用）
monthly_excel_for_write = monthly_pandas_data.reset_index()
# 创建xlwt的workbook
workbook = xlwt.Workbook(encoding='UTF-8')
# 添加sheet（名字为'空气温度（0.1℃）'）
worksheet = workbook.add_sheet(normal_data.replace(' ','')) 


# ## Input Headers

# In[8]:


## 日期
### 设置初始行为0
cur_row = 0
### 在第0行写入日期（2019年02月）
worksheet.write_merge(cur_row,cur_row,2,4,str(year)+'年'+str(month)+'月',font_style_lrtb_0)

## 一级表头
### 增加一行
cur_row += 1
### 使用write_merge分别写入第一行的表头内容
### 使用辅助程序按需求对单元格进行定义
worksheet.write_merge(cur_row,cur_row+1,0,0,'日期',font_style_lrtb)
worksheet.write_merge(cur_row,cur_row,1,24,normal_data,font_style_lrt)
worksheet.write_merge(cur_row,cur_row,25,26,avg_data,font_style_lrt)
worksheet.write_merge(cur_row,cur_row+1,27,27,max_data,font_style_lrtb)
worksheet.write_merge(cur_row,cur_row+1,28,28,min_data,font_style_lrtb)

## 二级表头
### 增加一行
cur_row += 1
### 使用write分别写入第二行的时间表头
### 使用辅助程序按需求对单元格进行定义
for i in range(1,25):
    if i in [6,12,18,24]:
        ### left right bottom 为粗框，其余为细框
        cur_style = font_style_lrb
    else:
        ### bottom 为粗框，其余为细框
        cur_style = font_style_b
    worksheet.write(cur_row,i,int(monthly_excel_for_write[normal_data].columns.values[i-1]),cur_style)
### 使用write分别写入第二行的剩余表头
### 使用辅助程序按需求对单元格进行定义
worksheet.write(cur_row,25,'4次',font_style_lb)
worksheet.write(cur_row,26,'24次',font_style_rb)


# # Input Index Column and Data

# In[9]:


## 自动录入数据
cur_row += 1

## 使用便利循环录入每一个数值
for i in range(len(monthly_excel_for_write)):
    for j in range(len(monthly_excel_for_write.iloc[i].values)):
        if i == len(monthly_excel_for_write)-1:  #最后一行，需要特别设置下粗框
            if j in [0,6,12,18,24]:              #设置第0，6，12，18，24列为左右下粗框
                cur_style = font_style_lrb
            elif j in [26,28]:                   #设置第26，28列为左右下粗框
                cur_style = font_style_rb
            else:                                #其余为正常细边框
                cur_style = font_style_b
        else:                                    #中间行，不需要下粗框
            if j in [0,6,12,18,24]:              #设置第0，6，12，18，24列为左右粗框
                cur_style = font_style_lr
            elif j in [26,28]:                   #设置第26，28列为左右下粗框
                cur_style = font_style_r
            else:                                #其余为正常细边框
                cur_style = font_style_normal
        cur_data = monthly_excel_for_write.iloc[i].values[j]      #获取data[i,j]数据
        worksheet.write(cur_row,j,_Int_Print(cur_data),cur_style) #使用_Int_Print录入表格
    ## 完成一行 增加一级
    cur_row += 1

for i in range(0,cur_row+1):
    worksheet.row(i).height_mismatch = True
    worksheet.row(i).height = 25*20         # 设置25行高 excel 1 行高 = 20 height
worksheet.col(0).width = 256 * 14           # 设置14列宽 excel 1 列宽 = 256 width


# In[10]:


workbook.save(r'.\Sample Result.xls')


# In[ ]:





# In[ ]:




