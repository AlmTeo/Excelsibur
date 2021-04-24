#!/usr/bin/env python3

import openpyxl, sys


def open_xlxs(given_path):
	xlsx = openpyxl.load_workbook(given_path)
	try:
		sheet = xlsx['template']
	except Exception as e:
		return "ERROR: Sheet name not found! Please rename the current sheet to: \"template\""
		


def generate_emails(first_name, last_name, mail):
	# Convert frst letter of string into lowercase
	no_upper = lambda test_str: test_str[:1].lower() + test_str[1:]

	first_name = no_upper(first_name)
	last_name = no_upper(last_name)


def find_names(no_of_inputs):
	# Iterate through all inputs and generate email from first name and last name
	for i in range(2, no_of_inputs + 2):
		first_name = sheet['O' + str(i)].value
		last_name = sheet['P' + str(i)].value
		found_status = sheet['M' + str(i)].value

		print(first_name)
		print(last_name)		
		print(found_status)




		if first_name == last_name == found_status == None:
			break



		# if first_name == None or last_name == None:
		# 	continue
		# else:
		# 	generate_emails(first_name, last_name, mail)


# xlsx.save(path)