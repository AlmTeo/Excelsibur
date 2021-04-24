#!/usr/bin/env python3
# from email_generator import
import PySimpleGUI as sg
import os 
from email_generator import *

def main_window():
	# Add a touch of color
	sg.theme('DarkPurple4')

	# All the stuff inside your main window.
	layout = [
				[sg.Text('Enter path of the excel file:'), sg.Input(), sg.FileBrowse()],
				[sg.Button('Ok'), sg.Button('Cancel')] 
			 ]

	# Create the Window
	window = sg.Window('Mail Generator', layout)

	# Event Loop to process "events" and get the "values" of the inputs
	while True:
		event, values = window.read()
		# If user closes window or clicks cancel
		if event in (sg.WIN_CLOSED, 'Cancel'): 
			break

		if event == 'Ok':
			# Check validity of introduced path
			file_check = os.path.isfile(values[0])
			if file_check == True:
				open_xlxs(values[0])
			else:
				invalid_path()


	window.close()







def invalid_path():
	sg.theme('DarkBrown4')

	layout = [  
				[sg.Text('Invalid path introduced!')],
				[sg.Button('Ok')] 
			 ]

	window = sg.Window('ERROR', layout,size=(215, 70))

	while True:  
		event, values = window.read()
		if event in (sg.WIN_CLOSED, 'Ok'):
			break
		if event == 'Ok':
			break

	window.close()













def another_window():
	# Add a touch of color
	sg.theme('DarkPurple3')

	layout = [  [sg.Text('Filename')],
				[sg.Input(), sg.FileBrowse()], 
				[sg.OK(), sg.Cancel()]] 


	# Create the Window
	window = sg.Window('New_Window', layout)

	# Event Loop to process "events"
	while True:             
		event, values = window.read()
		if event in (sg.WIN_CLOSED, 'Cancel'):
			break

	window.close()


main_window()