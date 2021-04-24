from openpyxl import Workbook
xlsx = Workbook()

# grab the active worksheet
worksheet = xlsx.active

# Data can be assigned directly to cells
worksheet['A1'] = 42

# Rows can also be appended
worksheet.append([1, 2, 3])

# Python types will automatically be converted
import datetime
worksheet['A2'] = datetime.datetime.now()

# Save the file
xlsx.save("sample.xlsx")