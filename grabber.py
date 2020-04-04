import requests
from threading import *
import pandas as pd
import xlrd,xlwt
import time

def get_reg_ids(filename):
    ids = []
    f = open(filename,"r")
    for line in f:
        ids.append(line.strip())
    f.close()
    return ids

def get_data(reg_id):
    print("Fetching %s data...."%(reg_id))
    #my_data = {REG_ID:XXX,ROLL_NO:XXX,NAME:XXX,GRADES=[],SGPA=XXX,STATUS=XXX}
    my_data = {'reg_id':"",'roll_no':"",'name':"",'grades':[],'sgpa':0.0,'status':""}
    df = pd.read_html(requests.get(base_url+reg_id).content)[-1]
    #print(df)
    my_data['name']=df[1][0]
    my_data['reg_id']=df[1][1]
    my_data['roll_no']=df[1][2]
    for i in range(5,17):
        my_data['grades'].append(df[1][i])
    my_data['sgpa']=df[0][17].split(" ")[2]
    my_data['status']="NA"
    return my_data


def async_get_data(reg_id):
    print("Fetching %s data...."%(reg_id))
    global data
    #my_data = {REG_ID:XXX,ROLL_NO:XXX,NAME:XXX,GRADES=[],SGPA=XXX,STATUS=XXX}
    my_data = {'reg_id':"",'roll_no':"",'name':"",'grades':[],'sgpa':0.0,'status':""}
    df = pd.read_html(requests.get(base_url+reg_id).content)[-1]
    #print(df)
    my_data['name']=df[1][0]
    my_data['reg_id']=df[1][1]
    my_data['roll_no']=df[1][2]
    for i in range(5,17):
        my_data['grades'].append(df[1][i])
    my_data['sgpa']=df[0][17].split(" ")[2]
    my_data['status']="NA"
    data[reg_id] = my_data
    #print(data[reg_id])

def set_header(sheet,content):
    sheet.write(0,0,"SNO")
    sheet.write(0,1,"REG ID")
    sheet.write(0,2,"ROLL NO")
    sheet.write(0,3,"NAME")
    i=4
    for _ in content['grades']:
        sheet.write(0,i,"SUBJECT-"+str(i-3))
        i+=1
    sheet.write(0,i,"SGPA")
    i+=1
    sheet.write(0,i,"STATUS")
    return sheet

def update_sheet(sheet,row,content):
    sheet.write(row,0,row)
    sheet.write(row,1,content['reg_id'])
    sheet.write(row,2,content['roll_no'])
    sheet.write(row,3,content['name'])
    i=4
    for grade in content['grades']:
        sheet.write(row,i,grade)
        i+=1
    sheet.write(row,i,content['sgpa'])
    i+=1
    sheet.write(row,i,content['status'])
    return sheet

t1 = time.time()

base_url = 'http://poornima.edu.in/results/result_btech_cloud_vsem.php?id='
regno_file = "reg_ids.txt"
result_file = "results.xls"
data={}
threads=[]
try:
    reg_ids = get_reg_ids(regno_file)
    for reg_id in reg_ids:
        t = Thread(target=async_get_data,args=(reg_id,))
        threads.append(t)
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    print("Total Entries found:",len(data.keys()))
    wb = xlwt.Workbook(result_file)
    result_sheet = wb.add_sheet("My Sheet")
    result_sheet = set_header(result_sheet,data[reg_ids[0]])
    row=0
    for reg_id in reg_ids:
        try:
            row+=1
            result_sheet = update_sheet(result_sheet,row,data[reg_id])
            wb.save(result_file)
        except:
            print("Error Saving ID:",reg_id)
    print("Results updated successfully!")
except Exception as e:
    print("Error Occored:",e)

t2=time.time()
print("It took %s seconds"%(str(t2-t1)))
