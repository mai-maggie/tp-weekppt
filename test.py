from openpyxl import Workbook, load_workbook

# make the worksheet for totality
wb = Workbook()
sheet_total = wb.active
sheet_total.title = 'total'

# insert column titles
column_titles = ['ASIN', '价格', '曝光', '点击', 'CTR', 'ACOS',
                 'ACOAS', '广告费',
                 '店铺销售额', '广告销售额', '店铺销量', '广告销量',
                 '毛利润', '毛利率', '排名']

column_titles_len = len(column_titles)
count = 0
for row in sheet_total.iter_rows(min_row=1,
                                 max_row=1, min_col=0,
                                 max_col=column_titles_len):
    for cell in row:
        if count < column_titles_len:
            cell.value = column_titles[count]
        count += 1

#clean up source 1 sheet
source1 = load_workbook('sample1.xlsx')
source_sheet1 = source1.active
source_sheet1.delete_rows(2,amount=1)


#make a function to copy from source1 to sheet_total
def copy_source1(source_row_num,dest_col):
    list_source=[]
    for i,row in enumerate(source_sheet1):
        if i==0:
            continue
        item=row[source_row_num].value
        list_source.append(item)
    list_len=len(list_source)
    i=0
    for row in sheet_total[f"{dest_col}2:{dest_col}{list_len+1}"]:
        for cell in row:
            cell.value=list_source[i]
            i+=1

#copy asin col=3
copy_source1(3,"A")

#copy 价格/竞价 col=7
copy_source1(7,"B")

#copy 曝光 col=9
copy_source1(9,"C")

#copy 点击 col=10
copy_source1(10,"D")

#copy CTR col=11
copy_source1(11,"E")

#copy ACOS col=17
copy_source1(17,"F")

#copy ACOAS col=18
copy_source1(18,"G")

#copy 广告花费 col=13
copy_source1(13,"H")

#copy 店铺销售额 col=14
copy_source1(14,"I")

#copy 广告销售额 col=14
copy_source1(15,"J")

#copy 店铺销量 col=23
copy_source1(23,"K")

#copy 广告销量 col=24
copy_source1(24,"L")


#clean up source 2 sheet
source2=load_workbook('sample2.xlsx')
source_sheet2 = source2.active
source_sheet2.delete_rows(1,amount=1)

#make a funtion to find index of asin
def find_index_source(lookup_value, lookup_col_num):
    source_list=[]
    for i,row in enumerate(source_sheet2):
        if i==0:
            continue
        item=row[lookup_col_num].value
        source_list.append(item)
    target_index=source_list.index(lookup_value)
    return target_index

#make a function to return the value of the look up value
def return_valve(target_index,return_row_num,dest_col):
    return_list=[]
    for i,row in enumerate(source_sheet2):
        if i==0:
            continue
        item=row[return_row_num].value
        return_list.append(item)
        return(return_list[target_index])










