# Import modules
import xlrd
import os 
import simplejson
import boto
import boto.s3.connection

access_key = 'put your access key here!'
secret_key = 'put your secret key here!'

conn = boto.connect_s3(
        aws_access_key_id = access_key,
        aws_secret_access_key = secret_key,
        host = 'objects.dreamhost.com',
        #is_secure=False,               # uncomment if you are not using ssl
        calling_format = boto.s3.connection.OrdinaryCallingFormat(),
       )


# Give the path of current directory
loc = os.getcwd()

# User input
filename = raw_input("Enter filename:")

# absolute path of file
loc = loc + "\\" + filename + ".xlsx"

# Open workbook 
try:
    with xlrd.open_workbook(loc) as book:
        #book = xlrd.open_workbook(loc)
        pass
except FileNotFoundError as e:
    print("FileNotFoundError:",e)
    sys.exit()
except OSError as e:
    print("OSError:", e)
    sys.exit()
except Exception as e:
    print(type(e), e) 
    sys.exit()

# Finding particular sheet
index = 0
for name in book.sheet_names():
    # if required sheet is found
    if name == "MICs List by CC":

       sheet = book.sheet_by_index(index)
       # Obtain total no. of rows for the sheet with data filled in 
       rows = sheet.nrows   

       # Obtain column names of 1st row
       keys = sheet.row_values(0) 

       # empty list to be filled in with Column names of first row    
       lst = []

       # Iterate through all rows
       for i in range(1,rows): 
           row = sheet.row_values(i)           
           count  = 0
           
           # Create an empty dictonary to hold key:value pairs of "Column Name:Column Value" for each row
           dict = {}
           for j in row:	      
              dict[keys[count]] = j
              count = count + 1
           # Append this dictonary as an element to the list  
           lst.append(dict)

    # Conter for next sheet          
    index = index + 1

# Closing workbook
book.release_resources()

# Releasing the resourse here
del book

# Creating a .json object
simplejson.dumps(lst)

# Creating buckets
from boto.s3.connection import Location

# THIS SEEMS TO BE GIVING AN ERROR SINCE PYTHON HAS MIGRATED TO boto3
#conn.create_bucket('mybucket', location=Location.EU)  