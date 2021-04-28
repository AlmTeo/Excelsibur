#!/usr/bin/env python3
import openpyxl, sys


# Get mail template
def get_mail(given_sheet, given_line):
	found_mail = given_sheet['S' + str(given_line)].value
	return found_mail


# Putting the generated mail in a specific cell
def update_mail(excel_file, given_path, given_sheet, given_line, given_letter, generated_mail):
	mail_cell = given_sheet[str(given_letter) + str(given_line)]
	mail_cell.value = generated_mail
	excel_file.save(given_path)


# Updating the mail status in a specific cell
def update_mail_status(excel_file, given_path, given_sheet, given_line, given_letter, mail_status):
	mail_status_cell = given_sheet[str(given_letter) + str(given_line)]
	if mail_status == 1:
		mail_status_cell.value = "Built based on company rule."
	if mail_status == 0:
		mail_status_cell.value = "Cannot generate email."

	excel_file.save(given_path)


# Get contact status
def get_contact_status(given_sheet, given_line):
	contact_status = given_sheet['M' + str(given_line)].value
	return contact_status


# Get names from specific line
def get_names(given_sheet, given_line):
	first_name = given_sheet['O' + str(given_line)].value
	last_name = given_sheet['P' + str(given_line)].value
	return first_name, last_name


# Check if name is valid for email generation
def check_names_validity(first_name, last_name):
	if '-' in first_name or '.' in first_name or '-' in last_name or '.' in last_name:
		return False
	else: 
		return True


# Generate the email based on first name, last name and mail termination
def generate_mail(first_name, last_name, mail_at):
	# Convert frst letter of string into lowercase
	no_upper = lambda given_str: given_str[:1].lower() + given_str[1:]

	first_name = no_upper(first_name)
	last_name = no_upper(last_name)

	mail = first_name + "." + last_name + mail_at
	return mail.replace(' ', '')


# Deletes the contents of a cell
def delete_cell_contents(excel_file, given_path, given_sheet, given_line, given_letter):
	delete_cell = given_sheet[str(given_letter) + str(given_line)]
	delete_cell.value = None
	excel_file.save(given_path)


def main(given_excel_path):
	given_path = given_excel_path

	# Open excel and target template
	xlsx = openpyxl.load_workbook(given_path)
	sheet = xlsx['template']
	current_line = 3
	count = 0

	while get_contact_status(sheet, current_line) != None:
		if get_mail(sheet, current_line) != None:		
			current_mail_at = get_mail(sheet, current_line)

		if get_contact_status(sheet, current_line) == "Unfound":
			delete_cell_contents(xlsx, given_path, sheet, current_line, 'S')

		if get_contact_status(sheet, current_line) == "Contact Discovered":
			current_names = get_names(sheet, current_line)
			if check_names_validity(current_names[0], current_names[1]) == True:
				generated_mail = generate_mail(current_names[0], current_names[1], current_mail_at)
				update_mail(xlsx, given_path, sheet, current_line, 'S', generated_mail)
				update_mail_status(xlsx, given_path, sheet, current_line, 'T', 1)
			else:
				delete_cell_contents(xlsx, given_path, sheet, current_line, 'S')
				update_mail_status(xlsx, given_path, sheet, current_line, 'T', 0)

		current_line += 1