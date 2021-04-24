#!/usr/bin/env python3
# from email_generator import
import PySimpleGUI as sg
import os 

# Add a touch of color
sg.theme('DarkAmber')   


# All the stuff inside your window.
layout = [
			[sg.Text('Enter path of the excel file:                          '), sg.InputText()],
			[sg.Text('Enter the number of names lines in the table:'), sg.InputText()],
			[sg.Button('Ok'), sg.Button('Cancel')] 
		 ]

def run_the_window():
	# Create the Window
	window = sg.Window('Mail Generator', layout)

	# Event Loop to process "events" and get the "values" of the inputs
	while True:
		event, values = window.read()
		# If user closes window or clicks cancel
		if event == sg.WIN_CLOSED or event == 'Cancel': 
			break

		if event == 'Ok':
			print('You entered ', values[0])
			print('You entered ', values[1])


			value = os.path.isfile(values[0])

	window.close()

run_the_window()