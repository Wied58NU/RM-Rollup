#!/anaconda3/bin/python

# https://automatetheboringstuff.com/chapter12/

# import load_workbook
from openpyxl import load_workbook


import os

scratch_dir = "/Users/jeffreywiedemann/Desktop/Resource_Planning/"
csvfile = open("/Users/jeffreywiedemann/Desktop/Resource_Planning/rollup.csv","w+")

for filename in os.listdir(scratch_dir):
   if filename.endswith(".xlsx"):

     #Uncomment below to see filename during testing. Testing?
     #print (scratch_dir + filename)

     # set file path
     #filepath="/Users/jeffreywiedemann/Desktop/Resource_Planning/jeff_for_python.xlsx"
     filepath = scratch_dir + filename
     
     wb=load_workbook(filepath)
     
     sheet = wb['Summary']
     
     A2=sheet['A2']
     A1=sheet['A1']
     
     print(A2.value, end=",")
     csvfile.write(A2.value + ",")


     print(A1.value, end=",")
     csvfile.write(A1.value + ",")

     sheet = wb['Meetings & Admin']

     M_B10=sheet['B10']

     print(M_B10.value)
     csvfile.write(str(M_B10.value) + ",")

     csvfile.write("\n")
