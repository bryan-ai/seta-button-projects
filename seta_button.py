import csv
import os
import sys
import subprocess
from pathlib import Path
import logging
import pygsheets
from time import sleep
import pandas as pd
import numpy as np

import files_helper 

WORKING_DIR = Path(os.getcwd())
GATHER_DIR = os.path.join(WORKING_DIR,"gather")
ZIP_DIR = os.path.join(WORKING_DIR,"zips")
PARENT_DIR = WORKING_DIR.parent.absolute()
TESTING_DIR = os.path.join(PARENT_DIR,"seta_button_test")
PAGES_OUTPUT_DIR = os.path.join(WORKING_DIR,"pages")
PAGES_COMPLETE_DIR = os.path.join(WORKING_DIR,"pages_complete")
TEMPLATES_DIR = os.path.join(WORKING_DIR,"page_templates")
SUBMISSIONS_DIR = os.path.join(WORKING_DIR,"submissions")
SUBMISSIONS_COPY_DIR = os.path.join(WORKING_DIR,"submissions_copy")
PROJECT_DIR = os.path.join(WORKING_DIR,"project_report_files")
FINAL_DIR = os.path.join(WORKING_DIR,"final")
EXTRACT_DIR = os.path.join(os.path.join(WORKING_DIR,"extract"))

CREDENTIALS_DIR = os.path.join(WORKING_DIR,"credentials")

PREDICT_SOURCE_SPREADSHEET = sys.argv[1]
STUDENT_LIST_WORKSHEET_REMOTE = "student_list"
TEAM_LIST_WORKSHEET_REMOTE = "team_list"
MARKS_SUMMARY_REMOTE = "[Imported] predict_marks_summary"
COMMENTS_REMOTE = "[Imported] project_comments"
SERVICE_FILE = os.path.join(CREDENTIALS_DIR,"bryan-athena-scraper-0c8981cd98b3.json")
#TODO make the students do this: build an uploader that asks for all the data, and pdf versions

#TODO Convert to separate tasks: First Extract and choose. Then clean and verify. Then convert. Then move PDFs

''' function authorise and open the connection to google sheets '''
def get_spreadsheet(spreadsheet):
	print(f"Authorising connection to {spreadsheet}")
	gc = pygsheets.authorize(service_file=SERVICE_FILE)
	sh = gc.open(spreadsheet)
	print(f"Opening {spreadsheet}")
	return sh

'''function opens the worksheet with name argument passed to it'''
def get_worksheet_by_title(sh, worksheet):
	'''select the section data sheet to read assessment name data'''
	print(f"Opening {worksheet}")
	ws = sh.worksheet('title',worksheet)
	return ws

'''function creates a dataframe from the specified worksheet object'''
def export_worksheet_to_dataframe(worksheet):
	print(f"exporting {worksheet} to dataframe")
	df = worksheet.get_as_df()
	return df

'''function to filter dataframe by any given value in any given column'''
def filter_dataframe_by_value(df, filter_col, filter_value):
	filtered_df = df[df[filter_col] == filter_value]
	return filtered_df

'''function to filter dataframe by any given value in any given column'''
def filter_dataframe_by_column(df, column_list):
	filtered_df = df[column_list]
	return filtered_df

def ipynb_to_string(template_path):
	
	with open(template_path, "r") as cover_ipynb:
		file_text = cover_ipynb.read()
	return file_text

def string_to_ipynb(string, output_path):
	with open(output_path, "w") as report_ipynb:
		report_ipynb.write(string)


'''Function to read cover page template notebook, mail-merge on {{}} and write to notebook'''
def create_cover_page(cover_dict={}, template_path=TEMPLATES_DIR, output_path=PAGES_OUTPUT_DIR):
	cover_page_string = ipynb_to_string(os.path.join(template_path,"cover_page_template.ipynb"))
	cover_page_string = cover_page_string.replace("{{team}}",cover_dict['team'])
	# print("Here we are!!",'\\n\\n'.join(cover_dict['members']))
	cover_page_string = cover_page_string.replace("{{members}}",cover_dict['members'])
	cover_page_string = cover_page_string.replace("{{project}}",cover_dict['project'])
	cover_page_string = cover_page_string.replace("{{mark}}",str(cover_dict['mark']))
	cover_page_filename = f"1{cover_dict['team']}_cover_page.ipynb"
	cover_page_path = os.path.join(output_path,cover_page_filename)
	string_to_ipynb(cover_page_string,cover_page_path)
	return cover_page_filename, cover_page_path

'''Function to read students marks template notebook, mail-merge on {{}} and write to notebook'''
def create_students_page(students_dict={}, template_path=TEMPLATES_DIR, output_path=PAGES_OUTPUT_DIR):
	students_page_string = ipynb_to_string(os.path.join(template_path,"students_page_template.ipynb"))
	students_page_string = students_page_string.replace("{{members}}",students_dict['members'])
	students_page_string = students_page_string.replace("{{Notes}}",students_dict['Notes'])
	students_page_filename = f"5{students_dict['team']}_students_page.ipynb"
	students_page_path = os.path.join(output_path,students_page_filename)
	string_to_ipynb(students_page_string,students_page_path)
	return students_page_filename, students_page_path

'''Function to read comments marks template notebook, mail-merge on {{}} and write to notebook'''
def create_comments_page(comments_dict={}, template_path=TEMPLATES_DIR, output_path=PAGES_OUTPUT_DIR):
	comments_page_string = ipynb_to_string(os.path.join(template_path,"comments_page_template.ipynb"))
	comments_page_string = comments_page_string.replace("{{Marker 1}}",comments_dict['Marker 1'])
	comments_page_string = comments_page_string.replace("{{Marker 2}}",comments_dict['Marker 2'])
	comments_page_string = comments_page_string.replace("{{Notebook}}",comments_dict['Notebook'])
	comments_page_string = comments_page_string.replace("{{Additional notes}}",comments_dict['Additional notes'])
	comments_page_string = comments_page_string.replace("{{App}}",comments_dict['App'])
	comments_page_filename = f"7{comments_dict['team']}_comments_page.ipynb"
	comments_page_path = os.path.join(output_path,comments_page_filename)
	string_to_ipynb(comments_page_string,comments_page_path)
	return comments_page_filename, comments_page_path

def dataframe_rows_to_string(df,team=""):
	df_string_list = []
	header_list = list(df)
	min_width = 0
	for header in header_list:
		if len(df_string_list) == 0:
			df_string_list = df[header].tolist()
		else:
			temp_string = df[header].tolist()
			for i in range(len(df_string_list)):
				df_string_list[i] = df_string_list[i] + " " + str(temp_string[i])
	
	for string in df_string_list:
		string_length = len(string)
		if string_length > min_width:
			min_width = string_length
	for i in range(len(df_string_list)):
		df_string_list[i]=df_string_list[i].replace("\"","\\\"")
		if "\n" in df_string_list[i]:
			df_string_list[i]=df_string_list[i].replace("\n","\\n\\n")
		if "\r" in df_string_list[i]:
			df_string_list[i]=df_string_list[i].replace("\r","\\n\\n")
		string_representation = "{0:<" + f"{min_width}" + "}"
		df_string_list[i] = string_representation.format(df_string_list[i])
	df_string = '\\n\\n'.join(df_string_list)
	return df_string

def rename_downloaded_submission_files(submittor_list=[],team="",download_path="submissions_copy"):
	for submittor in submittor_list:
		rename_string = f"mv {download_path}/{submittor}*.zip {download_path}/{submittor}_{team}.zip"
		logging.info(f"TERMINAL COMMAND: {rename_string}")
		subprocess.call(rename_string, shell=True)


def unzip_and_move(submittor_list=[],team="",download_path="submissions_copy", extract_path=EXTRACT_DIR):
	print(f"{team} submittor list is {submittor_list}")
	submissions = len(submittor_list)
	no_submissions = submissions == 0
	multiple_submissions = submissions > 1
	if no_submissions:
		print(f"WARNING: We cannot find a submittor for {team}. Please check your submission list to see if anyone submitted for {team}. Please check the dropped list to see if the submittor has since been dropped from the programme")
	if multiple_submissions:
		print(f"Team {team} has {submissions} submissions by students with IDs {submittor_list}")
		for submittor in submittor_list:
			#TODO At this time, the user must make a decision about zip files before they know what is inside them. find a way to allow them to keep peeking until the user knows which one(s) they want
			print(f"Here are {submittor}_{team}'s files")
			unzip_peek_string = f"unzip -l {download_path}/{submittor}_{team}.zip"
			logging.info(f"TERMINAL COMMAND: {unzip_peek_string}")
			subprocess.call(unzip_peek_string,shell=True)
		for submittor in submittor_list:
			carry_on = input(f"Do you want to include {submittor}_{team}'s files in the document? [Y]es or [N]o: ")
			if carry_on.lower() == "yes" or carry_on.lower() == 'y':
				print(f"including {submittor}_{team}")
				# Leave out the -o flag to give user options to skip, replace, or keep both files
				unzip_string = f"unzip -jq {download_path}/{submittor}_{team}.zip -d {extract_path}" 
				logging.info(f"TERMINAL COMMAND: {unzip_string}")
				subprocess.call(unzip_string, shell=True)
				continue
			else: continue
	else:
		print(f"{team} submittor is {submittor_list[0]}")
		submittor = submittor_list[0]
		unzip_string = f"unzip -joq {download_path}/{submittor}_{team}.zip -d {extract_path}" 
		logging.info(f"TERMINAL COMMAND: {unzip_string}")
		subprocess.call(unzip_string, shell=True)
		# print(unzip_string)

def verify_zips_in_submissions_folder():
	file_list = os.listdir(SUBMISSIONS_COPY_DIR)
	filetype_list = []
	for file in file_list:
		if len(file.split(".")) == 1:
			filetype_list.append("unknown")
		else:
			filetype_list.append(file.split(".")[1])
	filetype_set = set(filetype_list)
	if len(filetype_set)!=1:
		print(f"The submission directory contains {len(filetype_set)} filetypes: {filetype_set}. The directory should contain only zips or only one '.' in the filename. Please attend to these manually")
		open_finder=input(f"Would you like to open a finder window for the submission? [Y]es or [N]o: \n")
		if open_finder.lower() == "yes" or open_finder.lower() == 'y':
			folder_to_show = SUBMISSIONS_COPY_DIR
			print(f"opening {folder_to_show}")
			logging.info(f"TERMINAL COMMAND: open -R {folder_to_show}")
			subprocess.call(["open", "-R", folder_to_show])
			continue_string = input("When you have dealt with the files, press Enter to continue")

def populate_dataframes():
	'''Create '''
	try:
		class_spreadsheet_object = get_spreadsheet(PREDICT_SOURCE_SPREADSHEET)
	except Exception as e:
		logging.info(e)
		print("Did you forget to provide a spreadheet name as argument?")
	student_list_worksheet_object = get_worksheet_by_title(class_spreadsheet_object,STUDENT_LIST_WORKSHEET_REMOTE)
	teams_list_worksheet_object = get_worksheet_by_title(class_spreadsheet_object,TEAM_LIST_WORKSHEET_REMOTE)
	marks_summary_worksheet_object = get_worksheet_by_title(class_spreadsheet_object,MARKS_SUMMARY_REMOTE)
	comments_worksheet_object = get_worksheet_by_title(class_spreadsheet_object, COMMENTS_REMOTE)

	'''Create Dataframe from student list worksheet and remove rows labeled as Dropped or Faculty'''
	student_list_dataframe = export_worksheet_to_dataframe(student_list_worksheet_object)
	student_list_dataframe=student_list_dataframe[student_list_dataframe.ATHENA !="Dropped"]

	
	team_list_dataframe = export_worksheet_to_dataframe(teams_list_worksheet_object)
	marks_summary_dataframe = export_worksheet_to_dataframe(marks_summary_worksheet_object)
	comments_dataframe = export_worksheet_to_dataframe(comments_worksheet_object)

	'''create a list of team names from the non-empty rows'''
	df_teams_no_blanks = marks_summary_dataframe['Team'].replace('',np.nan).dropna()
	teams_list = df_teams_no_blanks.tolist()
	'''create dataframe of submittors'''
	submittor_df = filter_dataframe_by_value(student_list_dataframe,'Submitted',"Download")
	return student_list_dataframe, team_list_dataframe, marks_summary_dataframe, comments_dataframe, teams_list, submittor_df

def extract_and_choose_files(teams_list=[]):
	for team in teams_list:
		#TODO make the students do this: build an uploader that asks for all the data, and pdf versions
		print_team_header(team, process="Extracting submission files")
		pages_complete_team_dir = os.path.join(PAGES_COMPLETE_DIR,team)
		print("Making pages and pages_complete directory")
		logging.info(f"TERMINAL COMMAND: mkdir {pages_complete_team_dir}")
		subprocess.call(["mkdir", pages_complete_team_dir])
		
		'''make team directory'''
		print(f"making directory for {team}")
		team_files_pdf_path = os.path.join(PROJECT_DIR,team)
		team_files_extracted_path = os.path.join(EXTRACT_DIR,team)
		team_directory_string = "mkdir " + team_files_pdf_path
		logging.info(f"TERMINAL COMMAND: {team_directory_string}")
		subprocess.call(team_directory_string, shell=True)

		'''Make student and submission zips '''
		team_submittor_list = filter_dataframe_by_value(submittor_df,'Team',team)['USER ID'].tolist()
		rename_downloaded_submission_files(submittor_list=team_submittor_list,team=team)
		unzip_and_move(submittor_list=team_submittor_list,team=team, download_path=SUBMISSIONS_COPY_DIR,extract_path=team_files_extracted_path)
	return 0

def trim_and_identify_files(teams_list = []):
	for team in teams_list:
		print_team_header(team, process="Choose and label files")
		'''Get valid submissions file list'''
		team_files_extracted_path = os.path.join(EXTRACT_DIR,team)
		files_dictionary = files_helper.get_files(team=team,files_path=team_files_extracted_path)
		print(f"Cleaning {team}'s folder. Please see logs for details of files removed")
		filetypes_list = list(files_dictionary)
		for filetype in filetypes_list:
			for filename in files_dictionary[filetype]:
				if filename.split('.')[1].isupper():
					files_helper.lower_file_extension_case(team=team, filename=filename,input_path=team_files_extracted_path,output_path=team_files_extracted_path)
		
		'''Convert jpeg ot png'''
		filetypes_list = list(files_dictionary)
		for filetype in filetypes_list:
			if(filetype == "jpeg"):
				files_helper.jpeg_to_png(team=team,input_path=team_files_extracted_path,output_path=team_files_extracted_path)
			'''Check for Duplicates, clean, and label'''
			if((filetype != "png") and (len(files_dictionary[filetype])>1)) or ((filetype == "ipynb") and (len(files_dictionary[filetype])==1)):
				for filename in files_dictionary[filetype]:
					files_helper.multi_file_helper(team=team, filename=filename, filetype=filetype,input_path=team_files_extracted_path)
	return 0

def create_students_info_pages(teams_list=[]):
		for team in teams_list:
			print_team_header(team, process="Creating information pages")
			pages_output_path =os.path.join(PROJECT_DIR,team)

			'''Making cover page'''
			print(f"Extracting {team} rows into Cover page Dictionary")
			team_name_df = filter_dataframe_by_value(team_list_dataframe,'Team', team)
			students_df = filter_dataframe_by_value(student_list_dataframe,'Team', team)
			students_string = dataframe_rows_to_string(students_df[['Name']])
			team_files_extracted_path = os.path.join(EXTRACT_DIR,team)
			cover_dict = {"team": team, "project":team_name_df['Project'].iloc[0], "mark":team_name_df['Mark'].iloc[0], "members":students_string}
			print(f"Generating cover page for {team}")
			cover_page_filename, cover_page_path=create_cover_page(cover_dict)
			

			'''Making students page'''
			print(f"Extracting {team} rows into students page dictionary")
			team_df = filter_dataframe_by_value(team_list_dataframe,'Team', team)
			students_df = filter_dataframe_by_value(student_list_dataframe,'Team', team)
			individual_string = dataframe_rows_to_string(students_df[['Name', 'Mark', 'Problem']])
			comment_string = dataframe_rows_to_string(team_df[['Notes']], team)
			students_dict = {"team": team, "Notes": comment_string, "members":individual_string}
			print(f"Generating indivuals page for {team}")
			students_page_filename, students_page_path=create_students_page(students_dict)
			

			'''Making comments page'''
			#TODO make this column title agnostic?
			print(f"Extracting {team} rows into students page dictionary")
			comment_df = filter_dataframe_by_value(comments_dataframe,'Team', team)
			marker_one_string = dataframe_rows_to_string(comment_df[['Marker 1']],team)
			marker_two_string = dataframe_rows_to_string(comment_df[['Marker 2']],team)
			marker_notebook_string = dataframe_rows_to_string(comment_df[['Notebook']],team)
			marker_additional_string = dataframe_rows_to_string(comment_df[['Additional notes']],team)
			marker_app_string = dataframe_rows_to_string(comment_df[['App']],team)
			comments_dict = {"team": team, "Marker 1":marker_one_string, "Marker 2":marker_two_string,"Notebook":marker_notebook_string, "Additional notes":marker_additional_string, "App":marker_app_string}
			comments_page_filename, comments_page_path=create_comments_page(comments_dict)

			'''convert pages to html'''
			
			files_helper.ipynb_to_pdf(team=team,file_list=[cover_page_filename,students_page_filename,comments_page_filename],input_path=PAGES_OUTPUT_DIR, output_path=pages_output_path)
		


def convert_and_move_files(teams_list = []):
		# Finally, move all pdfs
		for team in teams_list:
			print_team_header(team, process="Converting")
			team_files_extracted_path = os.path.join(EXTRACT_DIR,team)
			team_files_pdf_path = os.path.join(PROJECT_DIR,team)
			'''Update valid submissions file lists to check if all files are now pdf'''
			files_dictionary = files_helper.get_files(team=team, files_path=team_files_extracted_path)
			filetypes_list = list(files_dictionary)

			for filetype in filetypes_list:
				print(f"Converting file to pdf. filetype is {filetype} from filestype_list")
				if filetype != "pdf":
					#TODO Figure out the input path for this section and for files_converter_helper
					files_helper.files_converter_helper(team=team,filetype=filetype,file_list=files_dictionary[filetype],input_path=team_files_extracted_path,output_path=team_files_pdf_path)

def move_pdfs(teams_list=[]):
	# Finally, move all pdfs
	for team in teams_list:
		print_team_header(team, process="moving PDFs to report directory")
		team_files_extracted_path = os.path.join(EXTRACT_DIR,team)
		team_files_pdf_path = os.path.join(PROJECT_DIR,team)
		'''Update valid submissions file lists to check if all files are now pdf'''
		files_dictionary = files_helper.get_files(team=team, files_path=team_files_extracted_path)
		filetypes_list = list(files_dictionary)
		for filetype in filetypes_list:
			if filetype == "pdf":
				files_helper.pdf_mover(team=team,filetype=filetype,file_list=files_dictionary[filetype],input_path=team_files_extracted_path,output_path=team_files_pdf_path)



def print_team_header(team = "", process=""):
		print("-------------")
		print(f"TEAM: {team} {process}")
		print("-------------")

def verify_teams_have_submitted(teams_list = [],submittor_df = None):
	for team in teams_list:

		
		'''Make student and submission zips '''
		team_submittor_list = filter_dataframe_by_value(submittor_df,'Team',team)['USER ID'].tolist()
		submissions = len(team_submittor_list)
		no_submissions = submissions == 0
		if no_submissions:
			print(f"WARNING: We cannot find a submittor for {team}. Please check your submission list to see if anyone submitted for {team}. Please check the dropped list to see if the submittor has since been dropped from the programme")
			#TODO Add options
			while True:
				refresh = input("Would you like to refresh the data source and continue? [Y]es or [N]o?\n")
				match refresh.lower():
					case "y":
						student_list_dataframe, team_list_dataframe, marks_summary_dataframe, comments_dataframe, teams_list, submittor_df = populate_dataframes()
						break
					case "n":
						break
					case _:
						print("Please enter 'y' for to refresh data or 'n' to skip to the next team")


if __name__ == "__main__":
	student_list_dataframe, team_list_dataframe, marks_summary_dataframe, comments_dataframe, teams_list, submittor_df = populate_dataframes()

	verify_zips_in_submissions_folder()
	verify_teams_have_submitted(teams_list,submittor_df)
	
	while True:
		step_input = input("what would you like to do?\n0. Refresh the data source\n1. Create cover, comment and marks pages\n2. Extract submissions\n3. Trim and choose files\n4. Convert files to pdf\n5. Move pdfs to report directory\n6. All of the above\nQuit. End the programme\n")
		match step_input.lower():
			case "0":
				student_list_dataframe, team_list_dataframe, marks_summary_dataframe, comments_dataframe, teams_list, submittor_df = populate_dataframes()
				verify_teams_have_submitted(teams_list,submittor_df)
			case "1":
				create_students_info_pages(teams_list=teams_list)
			case "2":
				extract_and_choose_files(teams_list=teams_list)
			case "3":
				trim_and_identify_files(teams_list=teams_list)
			case "4":
				convert_and_move_files(teams_list=teams_list)

			case "5":
				move_pdfs(teams_list=teams_list)
			case "6":
				create_students_info_pages(teams_list=teams_list) 
				extract_and_choose_files(teams_list=teams_list)
				trim_and_identify_files(teams_list=teams_list)
				convert_and_move_files(teams_list=teams_list)
			case "quit":
				print("thank you! Goodbye") 
				break
			case _:
				print("please input a number from 0")
				continue



		
	
	# 	'''Update valid submissions file lists to check if all files are now pdf'''
	# 	files_dictionary = files_helper.get_files(team=team, files_path=team_files_extracted_path)
	# 	filetypes_list = list(files_dictionary)
	# 	for filetype in filetypes_list:
	# 		if filetype == "pdf":
	# 			files_helper.pdf_mover(team=team,filetype=filetype,file_list=files_dictionary[filetype],input_path=team_files_extracted_path,output_path=team_files_pdf_path)
					
		





		

