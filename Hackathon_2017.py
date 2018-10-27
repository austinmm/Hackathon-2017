import pandas
import xlwt
from datetime import datetime as dt

def get_sheet_name():
    name = input("Please enter an output file name ending in \".xls\", for example(file_name.xls): ")
    return name
     
def get_header_num(df):
    num = 0
    for i in df.columns:
        num += 1
    return num

def get_header_names(df, header_num):
    header_names = []
    for i in range(header_num):
        header_names.append(df.columns[i])
    return header_names

def get_occurances(sheet, column, column_num):
    value = []
    n = 0
    most_occurance = '0'
    most_occurance_name = ''
    least_occurance = '100'
    least_occurance_name = ''
    while n < len(column):
        a = column[n]
        if a not in value:
            occurances = column.count(a)
            if a != a:
                a = 0
            #sheet.write(6,column_num,occurances)
            value.append(a)
            if occurances > float(most_occurance):
                most_occurance = str(occurances)
                most_occurance_name = str(a)
            elif occurances < float(least_occurance):
                least_occurance = str(occurances)
                least_occurance_name = str(a)
        else:
            pass
        n += 1
    most_occurances = most_occurance_name + ": " + str(most_occurance) + "x"
    least_occurances = least_occurance_name + ": " + str(least_occurance) + "x"
    sheet.write(4,column_num,most_occurances)
    sheet.write(5,column_num,least_occurances)

def get_avg_value(sheet, column, column_num):
    if type(column[0]) is str:
        sheet.write(3,column_num,'--')

    else:
        avg_value = 0.0
        n = 0
        while n < len(column):
            a = column[n]
            if a != a:
                a = 0
            avg_value += a
            n += 1
        avg_value = str(int(avg_value / n))
        sheet.write(3,column_num,avg_value)

def get_max_min(sheet, column, column_num):
    if type(column[0]) is str:
        sheet.write(1,column_num,'--')
        sheet.write(2,column_num,'--')
    else:
        largest = max(column)
        smallest = min(column)
        if smallest != smallest:
            n = 0
            for i in column:
                if i == 0:
                    n = i
                else:
                    if i < n:
                        n = i
                    else:
                        pass
        sheet.write(1,column_num,str(largest))
        sheet.write(2,column_num,str(smallest))

def populate_column(i, Format, df):
        column_header = []
        for value in df[Format[i]].values:
               column_header.append(value)
        return column_header

def excel_setup(sheet):
    sheet.write(1,0,'Maximum Value')
    sheet.write(2,0,'Minimum value')
    sheet.write(3,0,'Average Value')
    sheet.write(4,0,'Most Occured Value')
    sheet.write(5,0,'Least Occured Value')
    sheet.write(6,0,'Newest Entry')
    sheet.write(7,0,'Oldest Entry')
    
def get_date(column_num ,sheet, column):
    n = 0
    oldest_date = "17-02-05"
    newest_date = "01-05-01"
    while n < len(column):
        a = str(column[n])
        a = a[2:10]
        if a != a:
            n += 1
            continue
        a_date = dt.strptime(a, "%y-%m-%d")
        old_date =  dt.strptime(oldest_date, "%y-%m-%d")
        new_date = dt.strptime(newest_date, "%y-%m-%d")
        if a_date > new_date:
            newest_date = a
        elif a_date < old_date:
            oldest_date = a
        n += 1
    for i in range(1,column_num):
        sheet.write(6,i,"--")
        sheet.write(7,i,"--")
    altered_new = str(newest_date)
    new_year = altered_new[:2]
    new_month = altered_new[3:5]
    new_day = altered_new[6:8]
    altered_old = str(oldest_date)
    old_year = altered_old[:2]
    old_month = altered_old[3:5]
    old_day = altered_old[6:8]
    sheet.write(6,column_num,"Year: 20%s, Month: %s, Day: %s" %(new_year,new_month,new_day))
    sheet.write(7,column_num,"Year: 20%s, Month: %s, Day: %s" %(old_year,old_month,old_day))
    
    
    
def main():
    excel_name = input("Please enter the name of your excel document (Case sensitive, make sure to add the extension .xlsx): ")
    sheet_name = get_sheet_name()
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Sheet_1')
    excel_setup(sheet)
    print("\nPlease wait while we process the excel data for you.\n")
    df = pandas.read_excel(excel_name)
    header_num = get_header_num(df)
    Format = get_header_names(df, header_num)
    n = 1
    for i in Format:
        sheet.write(0,n,i)
        n += 1   
    column_num = 1      
    for i in range(header_num):
        column = populate_column(i, Format, df)
        if "Date" in Format[i] or "date" in Format[i] and "/" in column[1]:
            get_date(column_num, sheet, column)
        else:
            get_occurances(sheet, column, column_num)
            get_avg_value(sheet, column, column_num)
            get_max_min(sheet, column, column_num)
            sheet.write(i+1,len(Format),"--")
        column_num += 1
    print("Your data has been processed and can be found in your excel file named %s" %(sheet_name))
    workbook.save(sheet_name)

main()
