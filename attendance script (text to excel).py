import numpy as np
import pandas as pd
import re
import csv
import os
from os import path

def getStudentNames(className):
    students = pd.read_excel(r'{}/{}th.xlsx'.format(className, className))
    df = students['NAME']
    dict_students = {}
    count=0
    for row in df:
        row = row.upper()
        dict_students[row] = 0
        count+=1

    return dict_students

def main():
    className = str(input("Attendance of class (9/10/11):  "))
    print(os.getcwd())
    os.chdir(".\Documents\Zoom")
    print(os.getcwd())
    students_dict = getStudentNames(className)
    
    folders = os.listdir('{}/'.format(className))
    for folder in folders:
        if(folder != '{}th.xlsx'.format(className) and folder != 'done' and folder != 'attendance {}th.csv'.format(className)) :
            # print(folder)
            file = '{}/meeting_saved_chat.txt'.format(folder)
            mark_attendance(className, students_dict, folder, file)
    print("Attendance Marked")
    os.remove("df.csv")
    os.remove("inter.csv")
    

def mark_attendance(className, stu_dict, folder, file):
   
    #creating list of names from dict
    names=[]
    for k,v in stu_dict.items():
        names.append(k)

    date = folder[:11]
    # print("folder: ", folder)
    # print("file: ", file)

    file = "./{}/{}/meeting_saved_chat.txt".format(className, folder)
    txtfile = open(file,"r", errors="ignore")

    # #reading the file and populating the data with list of lines
    data = txtfile.readlines()
    txtfile.close()

    # data = pd.read_csv(file, error_bad_lines=True,buffer_lines=None, memory_map=False)

    ### logic: using regular expression for searching name in the line and setting value for students who are present with 1
    for name in names:
        for line in data:
            # print(line)
            line=line.upper()
            if re.search(name, line):
                stu_dict[name] = 1


    ### exporting the data into a excel sheet

    #converting dict into df
    df = pd.DataFrame(list(stu_dict.items()),columns = ['name','present'])
    df.to_csv('df.csv')

    if(not path.exists("./{}/attendance {}th.csv".format(className, className))):
        df1 = df[['name']]
        df1.set_index("name", inplace=True)
        df1.to_csv("./{}/attendance {}th.csv".format(className, className)) 
        

    count=0
    with open("./{}/attendance {}th.csv".format(className, className)) as input, open('inter.csv', 'w', newline='') as output:
        writer = csv.writer(output)
        for row in csv.reader(input):
            if any(field.strip() for field in row):
                if count!=0:
                    row.append(stu_dict[row[0]])
                    print(row)
                else:
                    # dirName = os.path.dirname(file)
                    date = folder[:11]
                    row.append(date)
                    print(row)
                writer.writerow(row)
                count+=1
        input.close()
        output.close()

    with open('inter.csv') as input, open(("./{}/attendance {}th.csv".format(className, className)), 'w', newline='') as output:
        writer = csv.writer(output)
        for row in csv.reader(input):
            if any(field.strip() for field in row):
                writer.writerow(row)
        input.close()
        output.close()

main()