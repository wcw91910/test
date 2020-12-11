import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("NTU Coin-0555c96087e3.json", scope)

client = gspread.authorize(creds)

sheet = client.open("NTU Coin").sheet1  # Open the spreadhseet

data = sheet.get_all_records()  # Get a list of all records

# print(data)
# numRows = sheet.col_values(2)
# newIndex = len(numRows) + 1
# ls = [str(len(numRows) + 1), "aaa@gmail.com", "12345678", "hihihi", "0"]
# print(ls)
# sheet.append_row(ls, table_range="A{}".format(newIndex))
  # Insert the list as a row at index 4
# print(numRows)
# print(len(numRows))
mail = sheet.col_values(2)
print(mail)
index = mail.index("b09701153@gmail.com") + 1
print(index)

userInfo = sheet.row_values(index)

print(userInfo)
# sheet.update_cell(userInfo[0], 5, 100)
