#!/anaconda3/bin/python
import string

# https://automatetheboringstuff.com/chapter12/

# import load_workbook
from openpyxl import load_workbook


import os

teams = ['DAPS', 'DBA' , 'INF', 'RCI', 'COLLAB' ]

team  = input("What team do ya want to rollup? ")

if team.upper() not in teams:
   print ("Not a Valid Team!")
   print (teams)
   quit()


scratch_dir = "/Users/jeffreywiedemann/Desktop/Resource_Planning/Team_Reports/" + team.upper() + "/"
#csvfile = open("/Users/jeffreywiedemann/Desktop/Resource_Planning/daps_rollup.csv","w+")
csvfile = open("/Users/jeffreywiedemann/Desktop/Resource_Planning/" + team.lower() + "_rollup.csv","w+")


print("name,dept,Total_Time_Away,Total_General_Admin,Total_Managerial_Admin_Time,Total_Support,Total_Consulting,Total_Other,Annual_CI_Project_Hours,Annual_AS_Project_Hours,Annual_IT_SS_Project_Hours,Annual_ISO_Project_Hours,Annual_Schools_or_Depts_Project_Hours")

csvfile.write("name,dept,Total_Time_Away,Total_General_Admin,Total_Managerial_Admin_Time,Total_Support,Total_Consulting,Total_Other,Annual_CI_Project_Hours,Annual_AS_Project_Hours,Annual_IT_SS_Project_Hours,Annual_ISO_Project_Hours,Annual_Schools_or_Depts_Project_Hours\n")

for filename in os.listdir(scratch_dir):
   if filename.endswith(".xlsx"):

     #Uncomment below to see filename during testing. Testing?
     #print (scratch_dir + filename)

     # set file path
     #filepath="/Users/jeffreywiedemann/Desktop/Resource_Planning/jeff_for_python.xlsx"
     filepath = scratch_dir + filename
     
     wb=load_workbook(filepath, data_only=True)
     
     sheet = wb['Summary']
     
     Name = sheet['A2']
     Dept = sheet['A1']
     
     Total_Time_Away = sheet['C16']
     Total_General_Admin = sheet['C22']
     Total_Managerial_Admin_Time = sheet['C25']
     Total_Support = sheet['C33']
     Total_Consulting = sheet['C37']
     Total_Other = sheet['C41'] 

     Annual_CI_Project_Hours = sheet['C50']
     Annual_AS_Project_Hours = sheet['C51']
     Annual_IT_SS_Project_Hours = sheet['C52']
     Annual_ISO_Project_Hours = sheet['C53']
     Annual_Schools_or_Depts_Project_Hours = sheet['C54']




     print(Name.value, end=",")
     csvfile.write(Name.value + ",")


     print(Dept.value, end=",")
     csvfile.write(Dept.value + ",")

     print(Total_Time_Away.value, end=",")
     csvfile.write(str(Total_Time_Away.value) + ",")

     print(Total_General_Admin.value, end=",")
     csvfile.write(str(Total_General_Admin.value) + ",")

     print(Total_Managerial_Admin_Time.value, end=",")
     csvfile.write(str(Total_Managerial_Admin_Time.value) + ",")

     print(Total_Support.value, end=",")
     csvfile.write(str(Total_Support.value) + ",")

     print(Total_Consulting.value, end=",")
     csvfile.write(str(Total_Consulting.value) + ",")

     print(Total_Other.value, end=",")
     csvfile.write(str(Total_Other.value) + ",")


     print(Annual_CI_Project_Hours.value, end=",")
     csvfile.write(str(Annual_CI_Project_Hours.value) + ",")

     print(Annual_AS_Project_Hours.value, end=",")
     csvfile.write(str(Annual_AS_Project_Hours.value) + ",")

     print(Annual_IT_SS_Project_Hours.value, end=",")
     csvfile.write(str(Annual_IT_SS_Project_Hours.value) + ",")

     print(Annual_ISO_Project_Hours.value, end=",")
     csvfile.write(str(Annual_ISO_Project_Hours.value) + ",")

     print(Annual_Schools_or_Depts_Project_Hours.value)
     csvfile.write(str(Annual_Schools_or_Depts_Project_Hours.value))




#     sheet = wb['Meetings & Admin']
#
#     M_B10=sheet['B10']
#
#     print(M_B10.value)
#     csvfile.write(str(M_B10.value) + ",")

     csvfile.write("\n")
