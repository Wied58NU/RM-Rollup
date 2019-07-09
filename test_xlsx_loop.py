#!/anaconda3/bin/python

import os

scratch_dir = "/Users/jeffreywiedemann/Desktop/Resource_Planning/"

for filename in os.listdir(scratch_dir):
   if filename.endswith(".xlsx"):
     # sfilename = filename.split("_")
     print (scratch_dir + filename)

