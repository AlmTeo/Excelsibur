#!/usr/bin/env python3

import openpyxl, sys
path = r"C:\Users\almmo\Desktop\ProiectExPy\example.xlsx"
xlsx = openpyxl.load_workbook(path)

# Check for sheet
try:
	sheet = xlsx['template']
except Exception as e:
	print("ERROR: Sheet name not found! Please rename the current sheet to: \"template\"")
	sys.exit()


def generate_emails(first_name, last_name, mail):
	# Convert frst letter of string into lowercase
	no_upper = lambda test_str: test_str[:1].lower() + test_str[1:]

	first_name = no_upper(first_name)
	last_name = no_upper(last_name)


def find_names(no_of_inputs):
	# Iterate through all inpunts and generate email from first name and last name
	for i in range(2, no_of_inputs + 2):
		first_name = sheet['O' + str(i)].value
		last_name = sheet['P' + str(i)].value

		if first_name == None or last_name == None:
			continue
		else:
			generate_emails(first_name, last_name, mail)

# find_names(8)


val = input("Enter your value: ")


# sheet['B3'] = "Altceva"

# xlsx.save(path)