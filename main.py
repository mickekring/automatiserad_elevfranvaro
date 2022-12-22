
import sys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
from tinydb import TinyDB, Query
import yaml
import time
import shutil
import os
from glob import glob
from datetime import datetime, timedelta
import paramiko
import locale
import pymsteams
import pyautogui
import os.path


### DATABASE INIT 

# Creates a database and sets up two tables

db = TinyDB('db.json')
search_db = Query()
table_stats_year = db.table('PupilsYear')
table_stats_year_sections = db.table('PupilsSections')
table_stats_resources = db.table('FreeResources')

### Set locale
locale.setlocale(locale.LC_ALL, 'sv_SE.UTF-8')

COLOR_RED = "bg-danger"
COLOR_ORANGE = "bg-warning"
COLOR_GREEN = "bg-success"
COLOR_NONE = ""



def global_vars():

	global excel_to_list
	global excel_to_list_student
	
	global TODAYS_DATE
	
	global conf
	
	global NUMBER_OF_PUPILS_ALL
	global NUMBER_OF_PUPILS_F3
	global NUMBER_OF_PUPILS_46
	global NUMBER_OF_PUPILS_79

	global pupil_groups

	global pupil_groups_03
	global pupil_groups_46
	global pupil_groups_79

	global pupil_sections

	excel_to_list = []
	excel_to_list_student = []

	#TODAYS_DATE = "2022-03-25" ### For testing
	TODAYS_DATE = datetime.today().strftime('%Y-%m-%d')

	conf = yaml.load(open('config.yml'), Loader=yaml.FullLoader)

	### Teams url if using the notification part
	myTeamsMessage = pymsteams.connectorcard(conf['teams']['url'])

	NUMBER_OF_PUPILS_ALL = conf['number_of_pupils']['total']
	NUMBER_OF_PUPILS_F3 = conf['number_of_pupils']['c03']
	NUMBER_OF_PUPILS_46 = conf['number_of_pupils']['c46']
	NUMBER_OF_PUPILS_79 = conf['number_of_pupils']['c79']

	pupil_groups = conf['classes_and_sections']['pupil_groups']

	pupil_groups_03 = conf['classes_and_sections']['pupil_groups_03']
	pupil_groups_46 = conf['classes_and_sections']['pupil_groups_46']
	pupil_groups_79 = conf['classes_and_sections']['pupil_groups_79']

	pupil_sections = conf['classes_and_sections']['pupil_sections']



def download_and_save_excel_report():

	# DOWNLOADING EXCEL REPORT FROM FNS

	user = conf['user']['username']
	password = conf['user']['password']

	service = Service(conf['chromedriver']['path'])
	browser = webdriver.Chrome(service=service)
	

	print("\n---\nDOWNLOADING LATEST REPORT FROM FNS\n---\n")

	print("Öppnar FNS och loggar in.")
	browser.get("https://fnsservicesso1.stockholm.se/sso-ng/saml-2.0/authenticate?customer=https://login001.stockholm.se&targetsystem=Skola24Widget")
	WebDriverWait(browser, 5)

	newURL = browser.current_url
	newURLReplace = newURL.replace("https://", "")

	loginURL = ("https://" + user + ":" + password + "@" + newURLReplace)
	browser.get(loginURL)

	print("Öppnar Report Center.")
	browser.get("https://fns.stockholm.se/ng/portal/start/compilation/reports-center")
	time.sleep(4)

	print("Letar efter Frånvaroanmälan per elev och klickar på länken.")
	link = browser.find_element(By.XPATH, ("//h3[text()='Frånvaroanmälan per elev']"))   
	link.click()
	time.sleep(4)

	print("Letar efter Grundrapport frånvaroanmälan och klickar")
	link = browser.find_element(By.XPATH, ("//span[text()='Grundrapport frånvaroanmälan']"))   
	link.click()
	time.sleep(4)

	print("Väljer rollen Skoladmin.")
	link = browser.find_element(By.XPATH, ("//select[@class='w-p-role-selector w-mb0']/option[text()='Skoladmin']"))   
	link.click()
	time.sleep(4)

	iframe = browser.find_elements(By.TAG_NAME, 'iframe')[0]
	browser.switch_to.frame(iframe)

	browser.switch_to.frame("MainFrame")
	browser.switch_to.frame("ViewFrame")
	browser.switch_to.frame("LeftFrame")

	print("Letar efter - och bockar för att visa sekretessmarkerade elever.")
	link = browser.find_element(By.XPATH, ("//img[contains(@id,'ShowIntegrityCBImg')]"))
	link.click()
	time.sleep(4)

	print("Klistrar in alla klasser som ska kontrolleras.")
	link = browser.find_element(By.XPATH, ("//input[@name='SelectStudentControl1$GroupCombo']"))
	link.send_keys(conf['rpa']['list'])
	time.sleep(4)

	print("Klickar på Visa rapport.")
	link = browser.find_element(By.XPATH, ("//*[@id='ctl177_showReportButton2']"))
	link.click()
	time.sleep(4)

	browser.switch_to.default_content()
	iframe = browser.find_elements(By.TAG_NAME, 'iframe')[0]
	browser.switch_to.frame(iframe)
	browser.switch_to.frame("MainFrame")
	browser.switch_to.frame("ViewFrame")
	browser.switch_to.frame("ViewFrame")

	# Här måste klicka på "Visa ändå" in någonstans.
	#print("Letar efter Visa rapporten ändå och klickar")
	
	#link = browser.find_element(By.XPATH, ("//*[@id='OpenPdfLink']"))
	#link.click()
	#time.sleep(4)
              
	print("Väljer att ladda ned som excelfil.")
	link = browser.find_element(By.XPATH, ("//*[@id='UCContentType']"))
	link.click()
	time.sleep(4)

	browser.switch_to.default_content()
	iframe = browser.find_elements(By.TAG_NAME, 'iframe')[0]
	browser.switch_to.frame(iframe)
	browser.switch_to.frame("IIFrame0")

	link = browser.find_element(By.ID, ("r69"))
	link.click()

	time.sleep(4)

	browser.close()



def download_and_save_excel_report_students():

	# DOWNLOADING EXCEL REPORT FROM FNS

	user = conf['user']['username']
	password = conf['user']['password']

	service = Service(conf['chromedriver']['path'])
	browser = webdriver.Chrome(service=service)

	print("\n---\nDOWNLOADING LATEST REPORT FROM FNS\n---\n")

	browser.get("https://fnsservicesso1.stockholm.se/sso-ng/saml-2.0/authenticate?customer=https://login001.stockholm.se&targetsystem=Skola24Widget")
	WebDriverWait(browser, 5)

	newURL = browser.current_url
	newURLReplace = newURL.replace("https://", "")

	loginURL = ("https://" + user + ":" + password + "@" + newURLReplace)
	browser.get(loginURL)

	browser.get("https://fns.stockholm.se/ng/portal/start/compilation/reports-center")
	time.sleep(2)

	link = browser.find_element(By.XPATH, ("//h3[text()='Frånvaroanmälan per elev']"))   
	link.click()
	time.sleep(2)

	link = browser.find_element(By.XPATH, ("//span[text()='Grundrapport frånvaroanmälan för skola']"))   
	link.click()
	time.sleep(2)

	link = browser.find_element(By.XPATH, ("//select[@class='w-p-role-selector w-mb0']/option[text()='Skoladmin']"))   
	link.click()
	time.sleep(2)

	iframe = browser.find_elements(By.TAG_NAME, 'iframe')[0]
	browser.switch_to.frame(iframe)

	browser.switch_to.frame("MainFrame")
	browser.switch_to.frame("ViewFrame")
	browser.switch_to.frame("LeftFrame")

	link = browser.find_element(By.XPATH, ("//img[contains(@id,'ShowIntegrityCBImg')]"))
	link.click()
	time.sleep(2)

	link = browser.find_element(By.XPATH, ("//*[@id='ctl177_showReportButton2']"))
	link.click()
	time.sleep(2)

	browser.switch_to.default_content()
	iframe = browser.find_elements(By.TAG_NAME, 'iframe')[0]
	browser.switch_to.frame(iframe)
	browser.switch_to.frame("MainFrame")
	browser.switch_to.frame("ViewFrame")
	browser.switch_to.frame("ViewFrame")

	# Här måste klicka på "Visa ändå" in någonstans.
	print("Letar efter Visa rapporten ändå och klickar")
	#link = browser.find_element(By.XPATH, ("//span[text()='Visa rapporten']"))   
	link = browser.find_element(By.XPATH, ("//*[@id='OpenPdfLink']"))
	link.click()
	time.sleep(2)

	link = browser.find_element(By.XPATH, ("//*[@id='UCContentType']"))
	link.click()
	time.sleep(2)

	browser.switch_to.default_content()
	iframe = browser.find_elements(By.TAG_NAME, 'iframe')[0]
	browser.switch_to.frame(iframe)
	browser.switch_to.frame("IIFrame0")

	link = browser.find_element(By.ID, ("r69"))
	link.click()

	time.sleep(4)

	browser.close()



def move_rename_and_delete_excel_file():

	print("\nMOVING AND RENAMING EXCEL FILE\n")

	file_name = glob(conf['paths']['download'])[0]

	file_source = conf['excel_file']['file_source']

	file_destination = conf['excel_file']['file_destination']


	os.rename(file_name, file_source)
	print("File renamed to report.xls")
	shutil.move(file_source, file_destination)
	print("File moved to destination.")



def move_rename_and_delete_excel_file_students():

	print("\nMOVING AND RENAMING EXCEL FILE\n")

	file_name = glob(conf['paths']['download'])[0]
	
	file_source = conf['excel_file']['file_source_student']

	file_destination = conf['excel_file']['file_destination_student']


	os.rename(file_name, file_source)
	print("File renamed to report.xls")
	shutil.move(file_source, file_destination)
	print("File moved to destination.")



def read_excel_file():

	### READ EXCEL FILE

	# Reading the excel file from FNS. In a function so that it will be
	# able to re-read when a new excel file is generated.

	global df
	global total_rows


	df = pd.read_excel('report.xls')

	print()
	total_rows = len(df.index)
	print("NUMBER OF ROWS IN EXCEL FILE: " + str(total_rows))
	print()




def read_excel_to_list_dict():

	### READING EXCEL FILE AND STORING IN LIST

	# Looping through the first column and storing the results in the list.

	for x in range(0, total_rows):

		excel_to_list_student.append({"name": str(df['Frånvaroanmälan för skola'][x]), "datefrom": str(df['Unnamed: 2'][x]), "dateto": str(df['Unnamed: 4'][x]), "timefrom": str(df['Unnamed: 6'][x]), "timeto": str(df['Unnamed: 8'][x]), "reportedby": str(df['Unnamed: 10'][x])})



def read_excel_file_student():

	### READ EXCEL FILE

	# Reading the excel file from FNS. In a function so that it will be
	# able to re-read when a new excel file is generated.

	global df
	global total_rows


	df = pd.read_excel('report_students.xls')

	print()
	total_rows = len(df.index)
	print("NUMBER OF ROWS IN EXCEL FILE: " + str(total_rows))
	print()



def read_excel_to_list():

	### READING EXCEL FILE AND STORING IN LIST

	# Looping through the first column and storing the results in the list.

	for x in range(0, total_rows):	

		excel_to_list.append(str(df['Frånvaroanmälan för elev'][x]))



def get_date_from_excel_file():

	### GET DATE FROM EXCEL FILE

	# Looking for - and storing todays date (in the file) to variable.

	global date

	for x in excel_to_list:
		
		if "Från och med" in x:

			date = x
		
		else:

			pass

	date = date.replace("Från och med : ", "")
	print("DATE: " + date)
	print()



def clean_up_words_in_excel_dectionary():

	words_list_name =["nan", "Rapporten", "Från", "Till och", "Skola", "Årstaskolan", "den ", "Frånvaroanmälan", "Namn"]
	words_list_reportedby =["Frånvarande - ogiltig"]

	for item in excel_to_list_student:

		for word in words_list_name:

			if word in item['name']:

				excel_to_list_student.remove(item)
			else:
				pass


	for item in excel_to_list_student:

		for word in words_list_reportedby:

			if word in item['reportedby']:

				excel_to_list_student.remove(item)
			else:
				pass


	print("\n############\n")

	for item in excel_to_list_student:
		print(item)
	print()
	print()



def find_students():

	student_list = conf['resurs']['elever']

	count_students = 0

	for item in excel_to_list_student:

		for student in student_list:

			if (student['name']) == item['name']:

				if "nan" in item['timefrom']:

					if table_stats_resources.search((search_db.Date == item['dateto']) & (search_db.Class == student['class'])):
						print("Hit! Already in DB.")
					else:
						print("Nope. Not in DB.")

						print("Namn: " + item['name'] + " Datum: " + item['dateto'] + " | Klockan: " + item['timefrom'] + " - " + item['timeto'] + " | HELDAG | Klass: " + student['class'])
						store_student_absence(item['dateto'], student['class'], item['timefrom'], item['timeto'])
						print("Stored in DB.")

						myTeamsMessage.color("edd1e7")
						myTeamsMessage.title("<strong>" + day + " " + kl_nu + "</strong>")
						myTeamsMessage.color("edd1e7")
						myTeamsMessage.text("Elev i klass " + student['class'] + " frånvarande heldag idag. " + student['resource'] + " fri att omdisponera." )
						
						myTeamsMessage.send()

					count_students += 1

				else:
					print("Namn: " + item['name'] + " Datum: " + item['dateto'] + " | Klockan: " + item['timefrom'] + " - " + item['timeto'])

	print("Antal frånvarande heldag: " + str(count_students))

	with open("web/resurs.php", "w", encoding="utf-8") as f1:
		f1.write('<h4><strong>LEDIGA</strong> RESURSER<br /></h4><p>Antal: ' + str(count_students) + '</p>')



def clean_up_words_in_excel_file():

	### CLEAN UP OF EXCEL FILE

	# Cleaning the results from the excel file from words listed
	# in the config.yml file. Everything but the pupils name and its 
	# class should be left.

	words_to_eremove_from_list = conf['word_to_remove']['list']

	while "nan" in excel_to_list:
		
		excel_to_list.remove("nan")


	for words in words_to_eremove_from_list:

		for item in excel_to_list:

			if words in item:

				excel_to_list.remove(item)

			else:

				pass



def clean_up_single_names_in_excel_file():
	
	### CLEAN UP TOO LONG NAMES

	# Names that are too long are broken up, hence the first name is
	# on its own row which will be double counted if not removed.

	for item in excel_to_list:

		if item not in pupil_groups and "," not in item:
			
			excel_to_list.remove(item)

		else:
			
			pass



def remove_duplicates_in_list():

	### REMOVE DUPLICATES

	# For some reason, there are duplicate records of absence in the
	# excel file out of FNS. This removes duplicates.

	global excel_to_list_cleaned

	excel_to_list_cleaned = list(dict.fromkeys(excel_to_list))


def print_list():

	### PRINTING THE MAIN LIST

	# Not used anymore, but can be used to print the results of the
	# main list at different staget for debug.

	for x in excel_to_list_cleaned:

		if x in pupil_groups:

			print("\n")

		else:

			pass

		print(x)



def count_pupils_absent_per_class():

	### MAIN STATS AND COUNTING OF NUMBER OF PUPILS ABSENT

	local_pupil_groups = conf['classes_and_sections']['pupil_groups']

	# Creating lists to store the classes that have absence and their
	# index number for iteration.

	today_pupil_groups = []
	today_pupil_groups_index = []

	# Checking length of the list for iteration
	
	rows = len(excel_to_list_cleaned)

	# Looping through main list for classes and if found, beeing stored
	# in today_pupil_groups and today_pupil_group_index.
	# After that, the class is removed from the local_pupil_group so
	# that we can count and store a 0 in the database for the classes 
	# that have no absence.
	
	for x in excel_to_list_cleaned:

		if x in local_pupil_groups:

			today_pupil_groups.append(x)
			today_pupil_groups_index.append(excel_to_list_cleaned.index(x))
			local_pupil_groups.remove(x)

		else:

			pass

	# Printing some basic info

	print("GROUPS WITH ABSENCE TODAY: " + str(today_pupil_groups))
	print()
	print("GROUPS WITH ABSENCE TODAY INDEXES: " + str(today_pupil_groups_index))
	print()
	number_of_todays_groups = len(today_pupil_groups)
	print("NUMBER OF GROUPS TODAY: " + str(number_of_todays_groups) + " OF " + str(len(pupil_groups)))
	print()
	print("NUMBER OF ROWS AFTER CLEAN UP: " + str(rows))
	print()


	### SINGLE CLASSES AND SECTIONS COUNT

	# Counting number of pupils absent in the different classes and adding
	# the pupil_counter number to its respective list and storing it in
	# the database.

	pupil_03 = []
	pupil_46 = []
	pupil_79 = []


	counter = 1

	# Looping the index numbers of the list containing the indexes of the
	# groups with absence - in range from index number + 1 to exclude the group
	# name, up to the next index number.
	# The else is for the last group in which the range is up to the total
	# number of rows.

	for index in today_pupil_groups_index:

		pupil_counter = 0
		
		if counter < number_of_todays_groups:

			print()
			pupil_group = excel_to_list_cleaned[index]
			print(pupil_group)

			for y in range(index + 1, (today_pupil_groups_index[counter])):
				print(excel_to_list_cleaned[y])
				pupil_counter += 1

			print()
			print("Students absent: " + str(pupil_counter))
			print()

			counter += 1

		else:

			print()
			pupil_group = excel_to_list_cleaned[index]
			print(pupil_group)

			for y in range(index + 1, rows):
				print(excel_to_list_cleaned[y])
				pupil_counter += 1

			print()
			print("Students absent: " + str(pupil_counter))


		# Searching the database to see if the records are already stored. If not, store it.
		# If pupil_counter is changed, the records will be deleted and restored.


		if table_stats_year.search((search_db.Date == date) & (search_db.Class == pupil_group)):
			
			print(pupil_group + " is already stored in DB")
			in_db = table_stats_year.search((search_db.Date == date) & (search_db.Class == pupil_group))
			print("Stored in DB: " + str(in_db))

			in_db_absent = in_db[0]['Absentees']
			
			if in_db_absent != pupil_counter:

				print()
				print("----------------\n")
				print("UPDATING...")
				print("Removing records from DB...")
				remove_group_year_absence(date, pupil_group)

				print("Saving new records to DB...")
				store_group_year_absence(date, pupil_group, pupil_counter)

				in_db = table_stats_year.search((search_db.Date == date) & (search_db.Class == pupil_group) & (search_db.Absentees == pupil_counter))
				print("NEW records in DB: " + str(in_db))

			else:
				pass
			
		else:
			store_group_year_absence(date, pupil_group, pupil_counter)
			print(pupil_group + " added to DB.")

		#############################


		### Storing number of absent pupils in their respective section list

		print()			
		print("------")


		if pupil_group in pupil_groups_03:
			pupil_03.append(pupil_counter)

		elif pupil_group in pupil_groups_46:
			pupil_46.append(pupil_counter)

		else:
			pupil_79.append(pupil_counter)


	########################

	# SECTIONS
	# Stores the sum of all classes in the corresponding section.  

	print("\nSECTIONS ABSENCE\n")

	for section in pupil_sections:

		if table_stats_year_sections.search((search_db.Date == date) & (search_db.Section == section)):
			
			in_db = table_stats_year_sections.search((search_db.Date == date) & (search_db.Section == section))

			print(section + " : Already in DB\n")
			print("Stored in DB: " + str(in_db))

			remove_group_year_sections_absence(date, section)

			if section == "F-3":
				store_group_year_sections_absence(date, section, sum(pupil_03))
				print("Stored new records in db")

			elif section == "4-6":
				store_group_year_sections_absence(date, section, sum(pupil_46))
				print("Stored new records in db")

			else:
				store_group_year_sections_absence(date, section, sum(pupil_79))
				print("Stored new records in db")


			if section == "F-3":
				print(section + " absentees: " + str(sum(pupil_03)))
			
			elif section == "4-6":
				print(section + " absentees: " + str(sum(pupil_46)))
			
			else:
				print(section + " absentees: " + str(sum(pupil_79)))

			print()

		else:

			if section == "F-3":
				store_group_year_sections_absence(date, section, sum(pupil_03))

			elif section == "4-6":
				store_group_year_sections_absence(date, section, sum(pupil_46))

			else:
				store_group_year_sections_absence(date, section, sum(pupil_79))
			
			print(section + " : Added to DB\n")


	# HANDLING OF CLASSES WITH NO ABSENCE
	# Add records in database for classes who did not have any absence
	# TODO: Fix remove and add

	print("-------\n\nCLASSES WITH NO ABSENCE TODAY\n")

	print(local_pupil_groups)
	print()
	
	for x in local_pupil_groups:

		if table_stats_year.search((search_db.Date == date) & (search_db.Class == x)):
			print(x + " already in DB\n")

		else:
			store_group_year_absence(date, x, 0)
			print(x + " added to DB\n")



def store_group_year_absence(date, pupil_group, pupil_counter):

	table_stats_year.insert({'Date': date, 'Class': pupil_group, 'Absentees': pupil_counter})



def store_student_absence(Date, Class, StartTime, EndTime):

	table_stats_resources.insert({'Date': Date, 'Class': Class, 'StartTime': StartTime, 'EndTime': EndTime})



def remove_group_year_absence(date, pupil_group):

	table_stats_year.remove((search_db.Date == date) & (search_db.Class == pupil_group))



def remove_group_year_sections_absence(date, pupil_section):

	table_stats_year_sections.remove((search_db.Date == date) & (search_db.Section == pupil_section))



def store_group_year_sections_absence(date, pupil_section, pupil_counter):

	table_stats_year_sections.insert({'Date': date, 'Section': pupil_section, 'Absent': pupil_counter})



def stats_today():

	# QUICK STATS TODAY SECTIONS AND TOTAL 
	
	sections_to_check = conf['classes_and_sections']['pupil_sections']

	
	# Get previous day to exclude saturday and sunday.

	days_to_subtract = 1

	previuos_day = datetime.now() - timedelta(days_to_subtract)
	previuos_day = datetime.strftime(previuos_day, '%Y-%m-%d')

	check_date_of_previuos_day_in_db = table_stats_year_sections.search((search_db.Date == previuos_day))

	while check_date_of_previuos_day_in_db == []:
		days_to_subtract += 1
		previuos_day = datetime.now() - timedelta(days_to_subtract)
		previuos_day = datetime.strftime(previuos_day, '%Y-%m-%d')
		check_date_of_previuos_day_in_db = table_stats_year_sections.search((search_db.Date == previuos_day))


	print()
	print("------\n")
	print("QUICK STATS TODAY: " + TODAYS_DATE + "\n")
	print("Pupils total: " + str(NUMBER_OF_PUPILS_ALL))
	print("Pupils year F-3: " + str(NUMBER_OF_PUPILS_F3))
	print("Pupils year 4-6: " + str(NUMBER_OF_PUPILS_46))
	print("Pupils year 7-9: " + str(NUMBER_OF_PUPILS_79))
	print()

	
	# TODAY'S STATS

	number_of_pupils_absent = table_stats_year_sections.search((search_db.Date == TODAYS_DATE) & (search_db.Section == "F-3"))
	num_f3 = (number_of_pupils_absent[0]['Absent'])

	number_of_pupils_absent = table_stats_year_sections.search((search_db.Date == TODAYS_DATE) & (search_db.Section == "4-6"))
	num_46 = (number_of_pupils_absent[0]['Absent'])
	
	number_of_pupils_absent = table_stats_year_sections.search((search_db.Date == TODAYS_DATE) & (search_db.Section == "7-9"))
	num_79 = (number_of_pupils_absent[0]['Absent'])

	total_absence = num_f3 + num_46 + num_79

	total_absence_percentage = round(((total_absence / NUMBER_OF_PUPILS_ALL) * 100), 2)
	num_f3_percentage = round(((num_f3 / NUMBER_OF_PUPILS_F3) * 100), 2)
	num_46_percentage = round(((num_46 / NUMBER_OF_PUPILS_46) * 100), 2)
	num_79_percentage = round(((num_79 / NUMBER_OF_PUPILS_79) * 100), 2)

	log(log_date, log_time, day, "F-3 Absent", str(num_f3))
	log(log_date, log_time, day, "4-6 Absent", str(num_46))
	log(log_date, log_time, day, "7-9 Absent", str(num_79))
	log(log_date, log_time, day, "F-9 Absent", str(total_absence))

	
	# PREVIOUS DAY STATS

	number_of_pupils_absent_yesterday = table_stats_year_sections.search((search_db.Date == previuos_day) & (search_db.Section == "F-3"))
	num_f3_yesterday = (number_of_pupils_absent_yesterday[0]['Absent'])
	
	number_of_pupils_absent_yesterday = table_stats_year_sections.search((search_db.Date == previuos_day) & (search_db.Section == "4-6"))
	num_46_yesterday = (number_of_pupils_absent_yesterday[0]['Absent'])

	number_of_pupils_absent_yesterday = table_stats_year_sections.search((search_db.Date == previuos_day) & (search_db.Section == "7-9"))
	num_79_yesterday = (number_of_pupils_absent_yesterday[0]['Absent'])

	total_absence_yesterday = num_f3_yesterday + num_46_yesterday + num_79_yesterday

	total_absence_yesterday_percentage = round(((total_absence_yesterday / NUMBER_OF_PUPILS_ALL) * 100), 2)
	num_f3_yesterday_percentage = round(((num_f3_yesterday / NUMBER_OF_PUPILS_F3) * 100), 2)
	num_46_yesterday_percentage = round(((num_46_yesterday / NUMBER_OF_PUPILS_46) * 100), 2)
	num_79_yesterday_percentage = round(((num_79_yesterday / NUMBER_OF_PUPILS_79) * 100), 2)

	
	# DIFFERENCE IN NUMBERS AND PERCENTAGE POINTS

	difference_f9_numbers = total_absence - total_absence_yesterday
	difference_f3_numbers = num_f3 - num_f3_yesterday
	difference_46_numbers = num_46 - num_46_yesterday
	difference_79_numbers = num_79 - num_79_yesterday

	difference_f9_percentages = round((total_absence_percentage - total_absence_yesterday_percentage), 2)
	difference_f3_percentages = round((num_f3_percentage - num_f3_yesterday_percentage), 2)
	difference_46_percentages = round((num_46_percentage - num_46_yesterday_percentage), 2)
	difference_79_percentages = round((num_79_percentage - num_79_yesterday_percentage), 2)

	print("ABSENCE (PREVIOUS DAY)\n")

	print("F-9: " + str(total_absence) + " (" + str(difference_f9_numbers) + ")" + " | " + str(total_absence_percentage) + " %" + " (" + str(difference_f9_percentages) + " pp)\n")
	#print("Previuos day F-9: " + str(total_absence_yesterday) + " | " + str(total_absence_yesterday_percentage) + "%\n")
	
	print("F-3: " + str(num_f3) + " (" + str(difference_f3_numbers) + ")" + " | " + str(num_f3_percentage) + " %" + " (" + str(difference_f3_percentages) + " pp)")
	#print("Previuos day F-3: " + str(num_f3_yesterday) + " | " + str(num_f3_yesterday_percentage) + "%\n")
	
	print("4-6: " + str(num_46) + " (" + str(difference_46_numbers) + ")" + " | " + str(num_46_percentage) + " %" + " (" + str(difference_46_percentages) + " pp)")
	#print("Previuos day 4-6: " + str(num_46_yesterday) + " | " + str(num_46_yesterday_percentage) + "%\n")
	
	print("7-9: " + str(num_79) + " (" + str(difference_79_numbers) + ")" + " | " + str(num_79_percentage) + " %" + " (" + str(difference_79_percentages) + " pp)")
	#print("Previuos day 7-9: " + str(num_79_yesterday) + " | " + str(num_79_yesterday_percentage) + "%\n")
	
	print("------\n")


	# HTML MODULES TODAY

	arrow_up = '<i class="far fa-arrow-alt-circle-up" style="color: #dc3545"></i>'
	arrow_down = '<i class="far fa-arrow-alt-circle-down" style="color: #7cb43f"></i>'
	arrow_right = '<i class="far fa-arrow-alt-circle-right" style="color: #17a2b8"></i>'

	
	# TODAY F-9 TOTAL

	if difference_f9_numbers < 0:
		arrow = arrow_down

	elif difference_f9_numbers > 0:
		arrow = arrow_up
		difference_f9_numbers = "+" + str(difference_f9_numbers)

	else:
		arrow = arrow_right

	html_total_today = '<div class="divbox"><p class="boxtitle"><strong>' + arrow + ' IDAG</strong> TOTALT</p><p class="boxnumber">' + str(total_absence) + '<span class="yesterday">' + " " + str(difference_f9_numbers) + "" + '</span></p><p class="tim">' + str(total_absence_percentage) + " % (" + str(difference_f9_percentages) + " pp)" + '</p><p class="infotext">Antal elever frånvarande idag och skillnaden mot igår</p></div>'

	with open("web/total_today.php", "w", encoding="utf-8") as f2:
			f2.write(html_total_today)


	# TODAY F-3

	if difference_f3_numbers < 0:
		arrow = arrow_down

	elif difference_f3_numbers > 0:
		arrow = arrow_up
		difference_f3_numbers = "+" + str(difference_f3_numbers)

	else:
		arrow = arrow_right

	html_total_today = '<div class="divbox"><p class="boxtitle"><strong>' + arrow + ' IDAG</strong> ÅK F-3</p><p class="boxnumber">' + str(num_f3) + '<span class="yesterday">' + " " + str(difference_f3_numbers) + "" + '</span></p><p class="tim">' + str(num_f3_percentage) + " % (" + str(difference_f3_percentages) + " pp)" + '</p><p class="infotext">Antal elever frånvarande idag och skillnaden mot igår</p></div>'


	# TODAY 4-6

	with open("web/f3_today.php", "w", encoding="utf-8") as f2:
			f2.write(html_total_today)

	if difference_46_numbers < 0:
		arrow = arrow_down

	elif difference_46_numbers > 0:
		arrow = arrow_up
		difference_46_numbers = "+" + str(difference_46_numbers)

	else:
		arrow = arrow_right

	html_total_today = '<div class="divbox"><p class="boxtitle"><strong>' + arrow + ' IDAG</strong> ÅK 4-6</p><p class="boxnumber">' + str(num_46) + '<span class="yesterday">' + " " + str(difference_46_numbers) + "" + '</span></p><p class="tim">' + str(num_46_percentage) + " % (" + str(difference_46_percentages) + " pp)" + '</p><p class="infotext">Antal elever frånvarande idag och skillnaden mot igår</p></div>'

	with open("web/46_today.php", "w", encoding="utf-8") as f2:
			f2.write(html_total_today)


	# TODAY 7-9

	if difference_79_numbers < 0:
		arrow = arrow_down

	elif difference_79_numbers > 0:
		arrow = arrow_up
		difference_79_numbers = "+" + str(difference_79_numbers)

	else:
		arrow = arrow_right

	html_total_today = '<div class="divbox"><p class="boxtitle"><strong>' + arrow + ' IDAG</strong> ÅK 7-9</p><p class="boxnumber">' + str(num_79) + '<span class="yesterday">' + " " + str(difference_79_numbers) + "" + '</span></p><p class="tim">' + str(num_79_percentage) + " % (" + str(difference_79_percentages) + " pp)" + '</p><p class="infotext">Antal elever frånvarande idag och skillnaden mot igår</p></div>'

	with open("web/79_today.php", "w", encoding="utf-8") as f2:
			f2.write(html_total_today)

	
	# QUICK STATS TODAY PER CLASS - HIGH TO LOW

	print("TODAYS ABSENCE PER CLASS - SORTED FROM HIGH TO LOW\n")

	today_year_list = []

	for x in conf['classes_and_sections']['pupil_groups_set']:
		
		number_of_pupils_in_x = conf['number_of_pupils'][x]

		number_of_pupils_absent = table_stats_year.search((search_db.Date == TODAYS_DATE) & (search_db.Class == x))
		pupil_number = (number_of_pupils_absent[0]['Absentees'])
		
		percentage_absent = round(((pupil_number / number_of_pupils_in_x) * 100), 2)

		list_item = {'Class': x, 'Absent': pupil_number, 'Pupils': number_of_pupils_in_x, 'Percentage': percentage_absent}
		today_year_list.append(list_item)

	today_year_list.sort(key=lambda x: x.get('Percentage'), reverse = True)

	for x in today_year_list:
		print(x['Class'] + ": " + str(x['Absent']) + " (of " + str(x['Pupils']) + ") | " + str(x['Percentage']) + " %")


	# QUICK STATS TODAY PER CLASS - 0-9

	print("\nTODAYS ABSENCE PER CLASS - SORTED FROM 0 - 9\n")

	today_year_list = []

	for x in conf['classes_and_sections']['pupil_groups_set']:

		number_of_pupils_in_x = conf['number_of_pupils'][x]

		number_of_pupils_absent = table_stats_year.search((search_db.Date == TODAYS_DATE) & (search_db.Class == x))
		pupil_number = (number_of_pupils_absent[0]['Absentees'])

		percentage_absent = round(((pupil_number / number_of_pupils_in_x) * 100), 2)

		list_item = {'Class': x, 'Absent': pupil_number, 'Pupils': number_of_pupils_in_x, 'Percentage': percentage_absent}
		today_year_list.append(list_item)


	for x in today_year_list:
		print(x['Class'] + ": " + str(x['Absent']) + " (of " + str(x['Pupils']) + ") | " + str(x['Percentage']) + " %")


	# MODULES
	# Classes with LOW absence today for html module

	year_absent_today = []
	year_absent_yesterday = []

	for x in conf['classes_and_sections']['pupil_groups_set']:

		year_without_absence = table_stats_year.search((search_db.Date == TODAYS_DATE) & (search_db.Class == x))
		
		year_absence = (year_without_absence[0]['Absentees'])
		year = (year_without_absence[0]['Class'])
		
		if year_absence < 3:
			
			year_absent_today.append(year)

		else:
			pass 


	year_absent_today_len = len(year_absent_today)

	for x in conf['classes_and_sections']['pupil_groups_set']:

		year_without_absence = table_stats_year.search((search_db.Date == previuos_day) & (search_db.Class == x))
		
		year_absence = (year_without_absence[0]['Absentees'])
		year = (year_without_absence[0]['Class'])
		
		if year_absence < 3:
			
			year_absent_yesterday.append(year)

		else:
			pass 

	year_absent_yesterday_len = len(year_absent_yesterday)
	year_absent_today_len = len(year_absent_today)

	year_absent_today = str(year_absent_today)
	year_absent_today = year_absent_today.replace("[", "")
	year_absent_today = year_absent_today.replace("]", "")
	year_absent_today = year_absent_today.replace("'", "")

	arrow_down_alt = '<i class="far fa-arrow-alt-circle-down" style="color: #dc3545"></i>'
	arrow_up_alt = '<i class="far fa-arrow-alt-circle-up" style="color: #7cb43f"></i>'

	
	difference_absent_years = year_absent_today_len - year_absent_yesterday_len

	if difference_absent_years > 0:
		thumbs_arrow = arrow_up_alt
		msg_down = " och antalet klasser med låg frånvaro har ökat sedan igår."
		difference_absent_years = ("+" + str(difference_absent_years))

	elif difference_absent_years < 0:
		thumbs_arrow = arrow_down_alt
		msg_down = ", men antalet klasser med låg frånvaro har minskat sedan igår."

	else:
		thumbs_arrow = arrow_right
		msg_down = ". Samma antal som igår."


	html_total_today = '<div class="divbox"><p class="boxtitle"><strong>' + thumbs_arrow + ' LÅG</strong> frånvaro</p><p class="boxnumber">' + str(year_absent_today_len) + '<span class="yesterday">' + " " + str(difference_absent_years) + "" + '</span></p><p class="tim">' + year_absent_today + '</p><p class="infotext">har max 2 elever frånvarande idag, normal frånvaro -' + msg_down + '</p></div>'

	with open("web/class_no_absence_today.php", "w", encoding="utf-8") as f2:
			f2.write(html_total_today)


	#M MODULES
	# Classes with HIGH absence today for html module

	year_absent_today = []
	year_absent_yesterday = []

	
	for x in conf['classes_and_sections']['pupil_groups_set']:

		year_with_high_absence = table_stats_year.search((search_db.Date == TODAYS_DATE) & (search_db.Class == x))
		
		year_absence = (year_with_high_absence[0]['Absentees'])
		year = (year_with_high_absence[0]['Class'])
		
		if year_absence > 4:
			
			year_absent_today.append(year)

		else:
			pass 


	year_absent_today_len = len(year_absent_today)

	for x in conf['classes_and_sections']['pupil_groups_set']:

		year_with_high_absence = table_stats_year.search((search_db.Date == previuos_day) & (search_db.Class == x))
		
		year_absence = (year_with_high_absence[0]['Absentees'])
		year = (year_with_high_absence[0]['Class'])
		
		if year_absence > 4:
			
			year_absent_yesterday.append(year)

		else:
			pass

	graph_high_absence = year_absent_today # To use for graph

	year_absent_yesterday_len = len(year_absent_yesterday)
	year_absent_today_len = len(year_absent_today)

	year_absent_today = str(year_absent_today)
	year_absent_today = year_absent_today.replace("[", "")
	year_absent_today = year_absent_today.replace("]", "")
	year_absent_today = year_absent_today.replace("'", "")

	arrow_down_alt = '<i class="far fa-arrow-alt-circle-up" style="color: #198754"></i>'
	arrow_up_alt = '<i class="far fa-arrow-alt-circle-down" style="color: #dc3545"></i>'

	
	difference_absent_years = year_absent_today_len - year_absent_yesterday_len

	if difference_absent_years < 0:
		thumbs_arrow = arrow_down
		msg_down = ", men antalet klasser med hög frånvaro har minskat sedan igår."

	elif difference_absent_years > 0:
		thumbs_arrow = arrow_up
		msg_down = " Tyvärr har vi fler klasser med hög frånvaro än igår."
		difference_absent_years = "+" + str(difference_absent_years)

	else:
		thumbs_arrow = arrow_right
		msg_down = ". Samma antal som igår."


	html_total_today = '<div class="divbox"><p class="boxtitle"><strong>' + thumbs_arrow + ' HÖG</strong> frånvaro</p><p class="boxnumber">' + str(year_absent_today_len) + '<span class="yesterday">' + " " + str(difference_absent_years) + "" + '</span></p><p class="tim">' + year_absent_today + '</p><p class="infotext">har 5 eller fler elever borta idag, över 15%' + msg_down + '</p></div>'

	with open("web/class_high_absence_today.php", "w", encoding="utf-8") as f2:
			f2.write(html_total_today)


	# MODULE PUPIL ABSENCE THIS YEAR - FROM HIGH TO LOW

	print("\nPUPIL ABSENCE THIS YEAR - FROM HIGH TO LOW\n")

	print("------\n")

	today_year_list = []

	for x in conf['classes_and_sections']['pupil_groups_set']:
		count = 0
		counting_days = []

		number_of_pupils_absent = table_stats_year.search((search_db.Class == x))
		
		for nums in number_of_pupils_absent:

			pupil_number = (nums['Absentees'])
			count += pupil_number

			count_days = (nums['Date'])
			counting_days.append(count_days)

		number_of_days = len(counting_days)
			
		today_year_list.append({'Class': x, 'Absent': count, 'Days': number_of_days})

	today_year_list.sort(key=lambda x: x.get('Absent'), reverse = True)

	html_list = []

	for x in today_year_list:

		average_absent_days = round((x['Absent']) / (x['Days']), 2)
		number_of_pupils_in_x = conf['number_of_pupils'][x['Class']]
		average_absense_percentage = round(((average_absent_days / number_of_pupils_in_x) * 100), 2)

		if average_absense_percentage > 14.99:
			color = COLOR_RED
		elif average_absense_percentage > 9.99:
			color = COLOR_ORANGE
		else:
			color = COLOR_GREEN

		
		print(x['Class'] + ": " + str(x['Absent']) + " | School days: " + str(x['Days']) + " | Average per day: " + str(average_absent_days) + " | " + str(average_absense_percentage) + " %")

		if average_absense_percentage < 9.99:
			average_absense_percentage_sort = ("0" + str(average_absense_percentage))
		else:
			average_absense_percentage_sort = str(average_absense_percentage)

		html_list.append('<p class="cattitle ' + average_absense_percentage_sort + '"><strong>' + x['Class'] + "</strong> | " + str(average_absent_days) + ' elever per dag</p><div class="progress"><div class="progress-bar ' + color + '" role="progressbar" style="width: ' + str(average_absense_percentage + 15) + '%" aria-valuenow="' + str(average_absense_percentage) + '" aria-valuemin="0" aria-valuemax="100">' + str(average_absense_percentage) + "%" + '</div></div>')
		html_list.sort(reverse=True)

	status_html_list = ("".join(html_list))

	with open("web/status.php", "w", encoding="utf-8") as f2:
		f2.write(status_html_list)



def stats_graphs_sections():
	
	print("\nCREATING GRAPHS!\n")

	# Query dates and number of dates from DB

	list_dates = []

	search_dates = table_stats_year_sections.search(search_db.Section == "F-3")

	for x in search_dates:
		list_dates.append(x['Date'])

	list_dates.sort()

	number_of_dates = len(list_dates)
	
	#print(list_dates)

	list_dates_html = str(list_dates)
	list_dates_html = list_dates_html.replace("[", "")
	list_dates_html = list_dates_html.replace("]", "")
	list_dates_html = list_dates_html.replace("'", '"')

	
	# Query number of absent pupils from list_date above

	list_f3 = []
	list_46 = []
	list_79 = []
	list_all = []

	
	search_section = table_stats_year_sections.search(search_db.Section == "F-3")
	search_section.sort(key=lambda x: x.get('Date'))

	for x in search_section:

		result = (x['Absent'])
		result = round(((result / NUMBER_OF_PUPILS_F3) * 100), 2)

		list_f3.append(result)

	
	search_section = table_stats_year_sections.search(search_db.Section == "4-6")
	search_section.sort(key=lambda x: x.get('Date'))

	for x in search_section:

		result = (x['Absent'])
		result = round(((result / NUMBER_OF_PUPILS_46) * 100), 2)

		list_46.append(result)


	search_section = table_stats_year_sections.search(search_db.Section == "7-9")
	search_section.sort(key=lambda x: x.get('Date'))

	for x in search_section:

		result = (x['Absent'])
		result = round(((result / NUMBER_OF_PUPILS_79) * 100), 2)

		list_79.append(result)

	
	for date in list_dates:

		#print("NEW DATE")

		sum_dates = 0

		search_section = table_stats_year_sections.search(search_db.Date == date)

		for x in search_section:

			#print(x)
			sum_dates = sum_dates + (x['Absent'])

		sum_dates = round(((sum_dates / NUMBER_OF_PUPILS_ALL) * 100), 2)
		list_all.append(sum_dates)


	list_f3_html = str(list_f3)
	list_f3_html = list_f3_html.replace("[", "")
	list_f3_html = list_f3_html.replace("]", "")
	list_f3_html = list_f3_html.replace("'", '"')

	#print(list_f3_html)

	list_46_html = str(list_46)
	list_46_html = list_46_html.replace("[", "")
	list_46_html = list_46_html.replace("]", "")
	list_46_html = list_46_html.replace("'", '"')

	#print(list_46_html)

	list_79_html = str(list_79)
	list_79_html = list_79_html.replace("[", "")
	list_79_html = list_79_html.replace("]", "")
	list_79_html = list_79_html.replace("'", '"')

	#print(list_79_html)

	list_all_html = str(list_all)
	list_all_html = list_all_html.replace("[", "")
	list_all_html = list_all_html.replace("]", "")
	list_all_html = list_all_html.replace("'", '"')


	html_graph = ('<script> const labels = [' + list_dates_html + ']; const NUMBER_CFG = {min: 0, max: 100}; const data = {labels: labels, datasets: [{ label: "Alla elever", data: [' 
		+ list_all_html + '], borderColor: "#bf2c2c", backgroundColor: "#bf2c2c",}, { label: "Årskurs F-3", data: [' 
		+ list_f3_html + '], borderColor: "#7db53f", backgroundColor: "#7db53f",}, { label: "Årskurs 4-6", data: [' 
		+ list_46_html + '], borderColor: "#ffd24c", backgroundColor: "#ffd24c",}, { label: "Årskurs 7-9", data: [' 
		+ list_79_html + '], borderColor: "#17a2b8", backgroundColor: "#17a2b8",}]}; const config = {type: "line", data: data, options: {color: "#ffffff", responsive: true, plugins: {legend: {position: "top",}, } }, }; var myChart = new Chart(document.getElementById("myChart30"), config);</script>')

	with open("web/total_and_sections.php", "w", encoding="utf-8") as f2:
			f2.write(html_graph)



def stats_graphs_last_15_days():
	
	print("\nCREATING GRAPH 2 - Last 15 days!\n")

	# Query dates and number of dates from DB

	list_dates = []
	latest_15_dates = []

	# Searching for all dates stored in DB
	search_dates = table_stats_year_sections.search(search_db.Section == "F-3")

	for x in search_dates:
		list_dates.append(x['Date'])

	list_dates.sort()

	number_of_dates = 15

	#print(list_dates)

	latest_15_dates = list_dates[-15:]

	print()
	print(latest_15_dates)

	list_dates_html = str(latest_15_dates)
	list_dates_html = list_dates_html.replace("[", "")
	list_dates_html = list_dates_html.replace("]", "")
	list_dates_html = list_dates_html.replace("'", '"')

	print()
	print(list_dates_html)
	print()

	# Query number of absent pupils from list_date above

	list_f3 = []
	list_46 = []
	list_79 = []
	list_all = []

	
	for x in latest_15_dates:

		print(x)
		searchdb = table_stats_year_sections.search((search_db.Section == "F-3") & (search_db.Date == x))
		print(searchdb)

		result = (searchdb[0]['Absent'])
		result = round(((result / NUMBER_OF_PUPILS_F3) * 100), 2)

		list_f3.append(result)


	for x in latest_15_dates:

		print(x)
		searchdb = table_stats_year_sections.search((search_db.Section == "4-6") & (search_db.Date == x))
		print(searchdb)

		result = (searchdb[0]['Absent'])
		result = round(((result / NUMBER_OF_PUPILS_46) * 100), 2)

		list_46.append(result)

	for x in latest_15_dates:

		print(x)
		searchdb = table_stats_year_sections.search((search_db.Section == "7-9") & (search_db.Date == x))
		print(searchdb)

		result = (searchdb[0]['Absent'])
		result = round(((result / NUMBER_OF_PUPILS_79) * 100), 2)

		list_79.append(result)

	
	for date in latest_15_dates:

		sum_dates = 0

		search_section = table_stats_year_sections.search(search_db.Date == date)

		for x in search_section:

			sum_dates = sum_dates + (x['Absent'])

		sum_dates = round(((sum_dates / NUMBER_OF_PUPILS_ALL) * 100), 2)
		list_all.append(sum_dates)


	list_f3_html = str(list_f3)
	list_f3_html = list_f3_html.replace("[", "")
	list_f3_html = list_f3_html.replace("]", "")
	list_f3_html = list_f3_html.replace("'", '"')

	print()
	print("LIST F-3")
	print(list_f3_html)
	print("------------")

	list_46_html = str(list_46)
	list_46_html = list_46_html.replace("[", "")
	list_46_html = list_46_html.replace("]", "")
	list_46_html = list_46_html.replace("'", '"')

	print()
	print("LIST 4-6")
	print(list_46_html)
	print("------------")

	#print(list_46_html)

	list_79_html = str(list_79)
	list_79_html = list_79_html.replace("[", "")
	list_79_html = list_79_html.replace("]", "")
	list_79_html = list_79_html.replace("'", '"')

	print()
	print("LIST 7-9")
	print(list_79_html)
	print("------------")

	#print(list_79_html)

	list_all_html = str(list_all)
	list_all_html = list_all_html.replace("[", "")
	list_all_html = list_all_html.replace("]", "")
	list_all_html = list_all_html.replace("'", '"')

	print()
	print("LIST ALL")
	print(list_all_html)
	print("------------")


	html_graph = ('<script> const labels2 = [' + list_dates_html + ']; const NUMBER_CFG2 = {min: 0, max: 100}; const data2 = {labels: labels2, datasets: [{ label: "Alla elever", data: [' 
		+ list_all_html + '], borderColor: "#bf2c2c", backgroundColor: "#bf2c2c",}, { label: "Årskurs F-3", data: [' 
		+ list_f3_html + '], borderColor: "#7db53f", backgroundColor: "#7db53f",}, { label: "Årskurs 4-6", data: [' 
		+ list_46_html + '], borderColor: "#ffd24c", backgroundColor: "#ffd24c",}, { label: "Årskurs 7-9", data: [' 
		+ list_79_html + '], borderColor: "#17a2b8", backgroundColor: "#17a2b8",}]}; const config2 = {type: "line", data: data2, options: {color: "#ffffff", responsive: true, plugins: {legend: {position: "top",}, } }, }; var myChart2 = new Chart(document.getElementById("myChart33"), config2);</script>')

	with open("web/graph_15_days.php", "w", encoding="utf-8") as f2:
			f2.write(html_graph)



def delete_excel_file():
	os.remove("report.xls")



def delete_excel_file_students():
	os.remove("report_students.xls")



def file_uploads_to_web():

	# Uploads php files to your sftp / web server
	
	#try:
	host = conf['sftp']['host']
	port = conf['sftp']['port']
	transport = paramiko.Transport((host, port))

	password = conf['sftp']['password']
	username = conf['sftp']['username']
	transport.connect(username = username, password = password)

	sftp = paramiko.SFTPClient.from_transport(transport)

	remoteUrlPath = conf['paths']['remoteurlpath']
	localUrlPath = conf['paths']['localurlpath']

	sftp.chdir(remoteUrlPath)

	filepath1 = "time.php"
	localpath1 = localUrlPath + "web/time.php"

	filepath2 = "status.php"
	localpath2 = localUrlPath + "web/status.php"

	filepath3 = "f3_today.php"
	localpath3 = localUrlPath + "web/f3_today.php"

	filepath4 = "46_today.php"
	localpath4 = localUrlPath + "web/46_today.php"

	filepath5 = "79_today.php"
	localpath5 = localUrlPath + "web/79_today.php"

	filepath6 = "total_and_sections.php"
	localpath6 = localUrlPath + "web/total_and_sections.php"

	filepath7 = "total_today.php"
	localpath7 = localUrlPath + "web/total_today.php"

	filepath8 = "class_no_absence_today.php"
	localpath8 = localUrlPath + "web/class_no_absence_today.php"

	filepath9 = "class_high_absence_today.php"
	localpath9 = localUrlPath + "web/class_high_absence_today.php"
	
	filepath10 = "base_info.php"
	localpath10 = localUrlPath + "web/base_info.php"

	filepath11 = "resurs.php"
	localpath11 = localUrlPath + "web/resurs.php"

	filepath12 = "all_classes.php"
	localpath12 = localUrlPath + "web/all_classes.php"

	filepath13 = "graph_15_days.php"
	localpath13 = localUrlPath + "web/graph_15_days.php"

	sftp.put(localpath1, filepath1)
	sftp.put(localpath2, filepath2)
	sftp.put(localpath3, filepath3)
	sftp.put(localpath4, filepath4)
	sftp.put(localpath5, filepath5)
	sftp.put(localpath6, filepath6)
	sftp.put(localpath7, filepath7)
	sftp.put(localpath8, filepath8)
	sftp.put(localpath9, filepath9)
	sftp.put(localpath10, filepath10)
	sftp.put(localpath11, filepath11)
	sftp.put(localpath12, filepath12)
	sftp.put(localpath13, filepath13)

	sftp.close()
	transport.close()

	print("\n>>> Files Successfully uploaded.")



def time_now():
	global kl_nu
	global day
	global date
	#global month
	global hour_now
	global log_date
	global log_time

	kl_nu = (time.strftime("%H:%M"))
	year = (time.strftime("%Y"))
	month = (time.strftime("%B"))
	date = (time.strftime("%d"))
	day = (time.strftime("%A"))
	hour_now = (time.strftime("%H"))

	log_date = (time.strftime('%Y-%m-%d'))
	log_time = (time.strftime("%H:%M"))

	print(kl_nu)
	print(day)
	print(hour_now)

	styleTime = ''
	version_number = conf['version']['number']
	version_date = conf['version']['date']

	with open("web/time.php", "w", encoding="utf-8") as f1:
		f1.write(styleTime + '<h1 class="clock"><i class="far fa-clock" aria-hidden="true"></i> ' 
			+ kl_nu + '</h1><h4><strong>SENAST</strong> UPPDATERAT<br />' + day + ' | ' + date + ' ' 
			+ month + ' ' + year + ' | ' + kl_nu + '</h4><p class="version"><i class="far fa-info-circle" aria-hidden="true"></i> Version ' + version_number + ' | ' + version_date + '</p>')



def stats_module_basic_info():

	pupils_total = conf['number_of_pupils']['total']
	pupils_f3 = conf['number_of_pupils']['c03']
	pupils_46 = conf['number_of_pupils']['c46']
	pupils_79 = conf['number_of_pupils']['c79']

	number_of_classes = 0
	
	for x in conf['classes_and_sections']['pupil_groups_set']:
		number_of_classes += 1

	with open("web/base_info.php", "w", encoding="utf-8") as f1:
		f1.write('<h4><strong>GRUND</strong>DATA</h4><p class="infotextleft">Antal elever totalt: ' + str(pupils_total) + '<br />Antal elever åk F-3: ' 
			+ str(pupils_f3) + '<br />Antal elever åk 4-6: ' + str(pupils_46) + '<br />Antal elever åk 7-9: ' + str(pupils_79) + '</p><p class="infotextleft">Antal klasser totalt: ' 
			+ str(number_of_classes) + '</p>')



def check_time_to_run():

	global go_live
	
	off_hours = ["14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "00", "01", "02", "03", "04", "05"]
	off_days = ["Lördag", "Söndag", "lördag", "söndag"]

	if hour_now in off_hours or day in off_days:
		print("Det är antingen eftermiddag eller helgdag.")
		go_live = 0

	else:
		print("Det är skoldag.")
		go_live = 1



def log(date, time, day, logtype, msg):
	
	with open("log.csv", "a") as error_log:

		error_log.write("\n" + date + "," + time + "," + day + "," + logtype + "," + msg + "")



def single_class_stats_module():

	color_orange = "#ff9007"
	color_red = "#dc3545"
	color_default = ""
	color_green = "#198754"

	list_class_absence = []
	
	#log_date_temp = "2022-03-16"
	#print(log_date)

	all_classes = conf['classes_and_sections']['pupil_groups_set']


	for x in all_classes:

		result = table_stats_year.search((search_db.Date == log_date) & (search_db.Class == x))
		#print(x)

		if result[0]['Absentees'] > 4:
			color = color_red
		elif result[0]['Absentees'] > 2:
			color = color_orange
		elif result[0]['Absentees'] == 0:
			color = color_green
		else:
			color = color_default

		list_class_absence.append({'Class': x, 'Absent': result[0]['Absentees'], 'Color': color})


	print(list_class_absence)

	html_f = '<div class="row modrow"><div class="col minimod" style="background: ' + list_class_absence[0]['Color'] + '"><p class="class-mod">0A</p><p class="num_mod">' + str(list_class_absence[0]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[1]['Color'] + '"><p class="class-mod">0B</p><p class="num_mod">' + str(list_class_absence[1]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[2]['Color'] + '"><p class="class-mod">0C</p><p class="num_mod">' + str(list_class_absence[2]['Absent']) + '</p></div><div class="col minimod" style="background: #11111100"></div><div class="col minimod" style="background: #11111100"></div></div>'
	
	html_1 = '<div class="row modrow"><div class="col minimod" style="background: ' + list_class_absence[3]['Color'] + '"><p class="class-mod">1A</p><p class="num_mod">' + str(list_class_absence[3]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[4]['Color'] + '"><p class="class-mod">1B</p><p class="num_mod">' + str(list_class_absence[4]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[5]['Color'] + '"><p class="class-mod">1C</p><p class="num_mod">' + str(list_class_absence[5]['Absent']) + '</p></div><div class="col minimod" style="background: #11111100"></div><div class="col minimod" style="background: #11111100"></div></div>'
	
	html_2 = '<div class="row modrow"><div class="col minimod" style="background: ' + list_class_absence[6]['Color'] + '"><p class="class-mod">2A</p><p class="num_mod">' + str(list_class_absence[6]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[7]['Color'] + '"><p class="class-mod">2B</p><p class="num_mod">' + str(list_class_absence[7]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[8]['Color'] + '"><p class="class-mod">2C</p><p class="num_mod">' + str(list_class_absence[8]['Absent']) + '</p></div><div class="col minimod" style="background: #11111100"></div><div class="col minimod" style="background: #11111100"></div></div>'

	html_3 = '<div class="row modrow"><div class="col minimod" style="background: ' + list_class_absence[9]['Color'] + '"><p class="class-mod">3A</p><p class="num_mod">' + str(list_class_absence[9]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[10]['Color'] + '"><p class="class-mod">3B</p><p class="num_mod">' + str(list_class_absence[10]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[11]['Color'] + '"><p class="class-mod">3C</p><p class="num_mod">' + str(list_class_absence[11]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[12]['Color'] + '"><p class="class-mod">3D</p><p class="num_mod">' + str(list_class_absence[12]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[13]['Color'] + '"><p class="class-mod">3E</p><p class="num_mod">' + str(list_class_absence[13]['Absent']) + '</p></div></div>'
		
	html_4 = '<div class="row modrow"><div class="col minimod" style="background: ' + list_class_absence[14]['Color'] + '"><p class="class-mod">4A</p><p class="num_mod">' + str(list_class_absence[14]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[15]['Color'] + '"><p class="class-mod">4B</p><p class="num_mod">' + str(list_class_absence[15]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[16]['Color'] + '"><p class="class-mod">4C</p><p class="num_mod">' + str(list_class_absence[16]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[17]['Color'] + '"><p class="class-mod">4D</p><p class="num_mod">' + str(list_class_absence[17]['Absent']) + '</p></div><div class="col minimod" style="background: #11111100"></div></div>'
	
	html_5 = '<div class="row modrow"><div class="col minimod" style="background: ' + list_class_absence[18]['Color'] + '"><p class="class-mod">5A</p><p class="num_mod">' + str(list_class_absence[18]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[19]['Color'] + '"><p class="class-mod">5B</p><p class="num_mod">' + str(list_class_absence[19]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[20]['Color'] + '"><p class="class-mod">5C</p><p class="num_mod">' + str(list_class_absence[20]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[21]['Color'] + '"><p class="class-mod">5D</p><p class="num_mod">' + str(list_class_absence[21]['Absent']) + '</p></div><div class="col minimod" style="background: #11111100"></div></div>'

	html_6 = '<div class="row modrow"><div class="col minimod" style="background: ' + list_class_absence[22]['Color'] + '"><p class="class-mod">6A</p><p class="num_mod">' + str(list_class_absence[22]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[23]['Color'] + '"><p class="class-mod">6B</p><p class="num_mod">' + str(list_class_absence[23]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[24]['Color'] + '"><p class="class-mod">6C</p><p class="num_mod">' + str(list_class_absence[24]['Absent']) + '</p></div><div class="col minimod" style="background: #11111100"></div><div class="col minimod" style="background: #11111100"></div></div>'
	
	html_7 = '<div class="row modrow"><div class="col minimod" style="background: ' + list_class_absence[25]['Color'] + '"><p class="class-mod">7A</p><p class="num_mod">' + str(list_class_absence[25]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[26]['Color'] + '"><p class="class-mod">7B</p><p class="num_mod">' + str(list_class_absence[26]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[27]['Color'] + '"><p class="class-mod">7C</p><p class="num_mod">' + str(list_class_absence[27]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[28]['Color'] + '"><p class="class-mod">7D</p><p class="num_mod">' + str(list_class_absence[28]['Absent']) + '</p></div><div class="col minimod" style="background: #11111100"></div></div>'
		
	html_8 = '<div class="row modrow"><div class="col minimod" style="background: ' + list_class_absence[29]['Color'] + '"><p class="class-mod">8A</p><p class="num_mod">' + str(list_class_absence[29]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[30]['Color'] + '"><p class="class-mod">8B</p><p class="num_mod">' + str(list_class_absence[30]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[31]['Color'] + '"><p class="class-mod">8C</p><p class="num_mod">' + str(list_class_absence[31]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[32]['Color'] + '"><p class="class-mod">8D</p><p class="num_mod">' + str(list_class_absence[32]['Absent']) + '</p></div><div class="col minimod" style="background: #11111100"></div></div>'
	
	html_9 = '<div class="row modrow"><div class="col minimod" style="background: ' + list_class_absence[33]['Color'] + '"><p class="class-mod">9A</p><p class="num_mod">' + str(list_class_absence[33]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[34]['Color'] + '"><p class="class-mod">9B</p><p class="num_mod">' + str(list_class_absence[34]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[35]['Color'] + '"><p class="class-mod">9C</p><p class="num_mod">' + str(list_class_absence[35]['Absent']) + '</p></div><div class="col minimod" style="background: ' + list_class_absence[36]['Color'] + '"><p class="class-mod">9D</p><p class="num_mod">' + str(list_class_absence[36]['Absent']) + '</p></div><div class="col minimod" style="background: #11111100"></div></div>'
	

	with open("web/all_classes.php", "w", encoding="utf-8") as f1:
		f1.write('<h4>ÅRSKURS F-3</h4>' + html_f + html_1 + html_2 + html_3 + '<h4>ÅRSKURS 4-6</h4>' + html_4 + html_5 + html_6 + '<h4>ÅRSKURS 7-9</h4>' + html_7 + html_8 + html_9)

### MAIN LOOP ###

def Main():

	while True:

		### GLOBAL VARIABLES THAT NEEDS RESET FOR LOOP TO FUNCTION

		global_vars()
		time_now()
		check_time_to_run()
		

		if go_live == 1: # Should be set to 1 for production

			error_msg = 0

			try:

				### RPA - GET REPORT FROM FNS
				
				download_and_save_excel_report()
				error_msg = 1

				move_rename_and_delete_excel_file()
				error_msg = 2
				
				read_excel_file()
				error_msg = 3

				read_excel_to_list()
				error_msg = 4

				get_date_from_excel_file()
				error_msg = 5

				clean_up_words_in_excel_file()
				error_msg = 6

				clean_up_single_names_in_excel_file()
				error_msg = 7

				remove_duplicates_in_list()
				error_msg = 8

				count_pupils_absent_per_class()
				error_msg = 9

				delete_excel_file()
				error_msg = 10
						
				### STATS

				stats_module_basic_info()
				error_msg = 11

				stats_today()
				error_msg = 12

				stats_graphs_sections()
				error_msg = 13

				stats_graphs_last_15_days()
				error_msg = 14

				single_class_stats_module()
				error_msg = 15
					
				file_uploads_to_web()
				error_msg = 16

				print("Success - part 1.")


				time.sleep(5)


				### PART II

				#check_if_file_exists = os.path.isfile('report_students.xls')

				#if check_if_file_exists == True:
				#	delete_excel_file_students()
				
				#else:
				#	pass

				#download_and_save_excel_report_students()
				#error_msg = 17

				#move_rename_and_delete_excel_file_students()
				#error_msg = 18

				#read_excel_file_student()
				#error_msg = 19

				#read_excel_to_list_dict()
				#error_msg = 20

				#clean_up_words_in_excel_dectionary()
				#error_msg = 21

				#find_students()
				#error_msg = 22
					
				#delete_excel_file_students()
				#error_msg = 23
					
				#file_uploads_to_web()
				#error_msg = 24

				#print("Success - part 2.")

				time.sleep(5)

				pyautogui.press('volumedown')
				time.sleep(60)
				pyautogui.press('volumeup')
				time.sleep(60)


			except:

				time.sleep(180)

				print("Error: " + str(error_msg))
				log(log_date, log_time, day, "Error: " + str(error_msg), "Main loop failed.")

		else:

			pyautogui.press('volumedown')
			time.sleep(60)
			pyautogui.press('volumeup')
			time.sleep(60)
			pass

		#log(log_date, log_time, day, "Sleep", "Main loop sleep for 10 min.")
		


### MAIN PROGRAM ###

if __name__ == "__main__":

	Main()

