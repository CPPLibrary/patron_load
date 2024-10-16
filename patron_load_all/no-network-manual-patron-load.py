import os
from os import path
import shutil
import csv
from datetime import date, timedelta
import glob
import time


def addNewPatron(p):
	"adds a new patron"
	today = date.today()
	nl = '\n'
	company = 'Staff'
	if 'student' in p[6]:
		company = 'Student'
	patron = 'I L2 U1 ^' + p[3] + '^^^' + nl
	patron += 'T Names' + nl
	patron += 'F FName ^' + p[1] + ' ' + p[2] + '^^^' + nl
	patron += 'F LName ^' + p[0] + '^^^' + nl
	patron += 'F Company ^' + company + '^^^' + nl
	patron += 'W' + nl
	patron += 'T UDF' + nl
	patron += 'F UdfNum ^1^^^' + nl
	patron += 'F UdfText ^' + p[3] + '^^^' + nl
	patron += 'W' + nl
	patron += 'T UDF' + nl
	patron += 'F UdfNum ^2^^^' + nl
	patron += 'F UdfText ^' + p[5] + '^^^' + nl
	patron += 'W' + nl
	patron += 'T Cards' + nl
	patron += 'F Code ^' + p[3][2:] + p[4] + '^^^' + nl
	patron += 'F CardNum ^' + p[5] + '^^^' + nl
	patron += 'F NumUses ^9999^^^' + nl
	patron += 'F StartDate ^' + today.strftime("%m/%d/%y") + '^^^' + nl
	patron += 'F StopDate ^' + '12/31/9999' + '^^^' + nl
	patron += 'F AddAcl ^Student Access^^^' + nl
	patron += 'W' + nl
	return patron

def updatePatronCard(p):
	"updates a patron card"
	nl = '\n'
	updatedPatron = 'I L2 U1 ^' + p[3] + '^^^' + nl
	updatedPatron += 'T Cards' + nl
	updatedPatron += 'F Code ^' + p[3][2:] + p[7] + '^^^' + nl
	updatedPatron += 'F ReplaceCode ^' + p[3][2:] + p[4] + '^^^' + nl
	updatedPatron += 'W' + nl
	return updatedPatron

def expirePatron(p):
	"expires a patron"
	nl = '\n'
	expiredPatron = 'I L2 U1 ^' + p[3] + '^^^' + nl
	expiredPatron += 'T Names' + nl
	expiredPatron += 'F FName ^' + p[1] + ' ' + p[2] + '^^^' + nl
	expiredPatron += 'F LName ^' + p[0] + '^^^' + nl
	expiredPatron += 'D' + nl
	return expiredPatron

def getLatestPatronLoad(f):
	"finds the latest added patron load"
	list_of_files = glob.glob(os.getcwd() + '\\' + f)
	if list_of_files:
		latest_file = max(list_of_files, key=os.path.getctime)
		return latest_file
	return ''

def processLoadingBar(n, c, d, l):
	"generates a loading bar and writes it to the screen"

	# n = num_patrons
	# c = counter
	# d = dot_count
	# l = last_checked
	# returns a dsx api entry for the patron

	c += 1
	percentage = c*100//n
	# print('\r' + str(n) + ' // ' + str(c) + ' // ' + str(d) + ' // ' + str(l) + ' // ' + str(percentage) + ' // ' + str(percentage != l), end='', flush=True)
	loading_bar = ''
	if(c == n):
		loading_bar = '\r100%  '
		dots = ''
		for d in range(50):
			dots += '.'
		loading_bar += '[' + str(dots).ljust(50) + ']\n'
		print('\r' + loading_bar, end='', flush=True)
	elif(percentage != l):
		l = percentage
		if(percentage%2 == 0 and percentage != 0):
			d += 1
		loading_bar += '\r' + str(percentage).rjust(3) + '%  '
		dots = ''
		for r in range(d):
			dots += '.'
		loading_bar += '[' + str(dots).ljust(50) + ']'
		print('\r' + loading_bar, end='', flush=True)
	return [c, d, l, loading_bar]

def printToScreenLog(l, t, p = True, w = True):
	"prints to both the log file and, optionally, the screen"
	if w:
		l.write(t + '\n')
	if p:
		print(t)

def process_exists(process_name):
	"checks if DSX is running"
	r = os.popen('tasklist /FI "imagename eq ' + process_name + '"').read().strip().split('\n')
	return r[-1].lower().startswith(process_name.lower())

def getLogFilename(f = '', i = 1):
	"gets the log filename; appends a sequential number if it exists"
	if(f == ''):
		f = 'logs\\' + date.today().strftime('libcard-%Y%m%d.log')
	replace = '.log'
	if(i != 1):
		replace = '-' + str(i - 1) + replace
	if(path.exists(f) == True):
		return getLogFilename(f.replace(replace, '-' + str(i) + '.log'), i = i + 1)
	else:
		return f

if(process_exists('DbSql.exe') == False):
	print('================================================================================')
	print('===                                                                          ===')
	print('===  Error! DSX is not running! Please start DSX and run this script again.  ===')
	print('===                                                                          ===')
	print('================================================================================')

	input("Press Enter to continue...")
	quit()

dsx_log_filename = getLogFilename()
dsx_log = open(dsx_log_filename, "w")
printToScreenLog(dsx_log, 'Created log at ' + os.getcwd() + '\\' + dsx_log_filename)

# Update the patron load dates manually
old_patron_load = getLatestPatronLoad('old-libcard-20240125.dat')
new_patron_load = getLatestPatronLoad('libcard-20240126.dat')
check_patron_load = old_patron_load.replace('old-', '')

if(check_patron_load == new_patron_load):
	# check if the latest patron load matches the previous patron load that was processed
	printToScreenLog(dsx_log, 'Previous patron load matches latest patron load. Aborting script.')
	printToScreenLog(dsx_log, '')
	if new_patron_load:
		os.remove(new_patron_load)
else:
	# latest patron load is new and needs to be processed
	printToScreenLog(dsx_log, 'Using ' + old_patron_load + ' for the previous patron load')
	printToScreenLog(dsx_log, 'Using ' + new_patron_load + ' for the latest patron load')
	dsx_host = 'C:\\Users\\CalPoly DSX client\\Desktop\\IMP Files'
	dsx_api = 'C:\\WinDSX\\API'
	# this is the file that will be uploaded to the DSX directory after processing the patron load here
	dsx_api_filename = '^IMP' + str(len(glob.glob(dsx_host + '\\*'))+1).zfill(2) + '.txt'

	start_time = time.monotonic()

	patrons_old = []
	patrons_exp = []
	patrons_new = []
	patrons_past = []
	patrons_newCard = 0
	patrons_updated = 0
	patrons_created = 0
	patrons_expired = 0
	patrons_past_num = 0
	patrons_current = 0

	printToScreenLog(dsx_log, '')
	# save the previous patron load to a list
	printToScreenLog(dsx_log, 'Saving data from previous patron load')
	if old_patron_load:
		with open(old_patron_load) as csvfile_old:
			readCSV_old = csv.reader(csvfile_old, delimiter=',')
			for p in readCSV_old:
				if p[6] == '|past_student|':
					patrons_past.append(p)
					patrons_past_num += 1
				else:
					patrons_old.append(p) # add each patron to patrons_old for checking
	else:
		printToScreenLog(dsx_log, 'No old patron load found!')

	# compare the previous and new patron loads
	printToScreenLog(dsx_log, '')
	if new_patron_load:
		# first pass gets the number of patrons in the patron load for the progress bar
		with open(new_patron_load) as csvfile_new:
			num_patrons = sum(1 for p in csvfile_new)

		# second pass processes the data
		with open(new_patron_load) as csvfile_new:
			readCSV_new = csv.reader(csvfile_new, delimiter=',')
			patrons_exp = patrons_old
			out = open(dsx_api_filename, "w")
			loading_bar = [0, 0, 0, '']
			counter = 0
			dot_count = 0
			last_checked = -1

			# process existing patrons
			printToScreenLog(dsx_log, 'Processing latest patron data')
			for p in readCSV_new:
				if p[6] == '|past_student|':
					patrons_past.append(p)
					patrons_past_num += 1
				else:
					patrons_new.append(p) # add each patron to patrons_new for checking
					patrons_current += 1
					for po in patrons_old:
						if po[5] == p[5]:
							if p[4] != po[4]: # patron has new card
								p[7] = po[4]
								out.write(updatePatronCard(p))
								patrons_newCard += 1
							#else: # some other data changed
								#out.write(addNewPatron(p))
								#patrons_updated += 1

							patrons_new.remove(p) # remove existing patrons from patrons_new
							patrons_exp.remove(po) # remove existing patrons from patrons_exp
							continue

				# show a loading indicator; process dots after the processes are done
				loading_bar = processLoadingBar(num_patrons, counter, dot_count, last_checked)
				counter = loading_bar[0]
				dot_count = loading_bar[1]
				last_checked = loading_bar[2]
			if loading_bar[3] != '':
				printToScreenLog(dsx_log, loading_bar[3], False)

			loading_bar = [0, 0, 0, '']
			counter = 0
			dot_count = 0
			last_checked = -1

			# process new patrons
			printToScreenLog(dsx_log, 'Processing new patrons')
			for p in patrons_new:
				out.write(addNewPatron(p))
				patrons_created += 1
				loading_bar = processLoadingBar(len(patrons_new), counter, dot_count, last_checked)
				counter = loading_bar[0]
				dot_count = loading_bar[1]
				last_checked = loading_bar[2]
			if loading_bar[3] != '':
				printToScreenLog(dsx_log, loading_bar[3], False)

			loading_bar = [0, 0, 0, '']
			counter = 0
			dot_count = 0
			last_checked = -1

			# process expired patrons
			printToScreenLog(dsx_log, 'Processing expired patrons')
			for p in patrons_exp:
				out.write(expirePatron(p))
				patrons_expired += 1
				loading_bar = processLoadingBar(len(patrons_exp), counter, dot_count, last_checked)
				counter = loading_bar[0]
				dot_count = loading_bar[1]
				last_checked = loading_bar[2]
			if loading_bar[3] != '':
				printToScreenLog(dsx_log, loading_bar[3], False)

			out.close()

		printToScreenLog(dsx_log, '')
		if(patrons_newCard == 0):
			printToScreenLog(dsx_log, 'No patrons received new cards')
		elif(patrons_newCard == 1):
			printToScreenLog(dsx_log, str(patrons_newCard) + ' patron received new cards')
		else:
			printToScreenLog(dsx_log, str(patrons_newCard) + ' patrons received new cards')

		if(patrons_updated == 0):
			printToScreenLog(dsx_log, 'No patrons were updated')
		elif(patrons_updated == 1):
			printToScreenLog(dsx_log, str(patrons_updated) + ' patron was updated')
		else:
			printToScreenLog(dsx_log, str(patrons_updated) + ' patrons were updated')

		if(patrons_created == 0):
			printToScreenLog(dsx_log, 'No patrons were added')
		elif(patrons_created == 1):
			printToScreenLog(dsx_log, str(patrons_created) + ' patron was added')
		else:
			printToScreenLog(dsx_log, str(patrons_created) + ' patrons were added')

		if(patrons_expired == 0):
			printToScreenLog(dsx_log, 'No patrons were expired')
		elif(patrons_expired == 1):
			printToScreenLog(dsx_log, str(patrons_expired) + ' patron was expired')
		else:
			printToScreenLog(dsx_log, str(patrons_expired) + ' patrons were expired')

		if(patrons_past_num > 0):
			printToScreenLog(dsx_log, str(patrons_past_num) + ' patrons are past student')

		if(patrons_current > 0):
			printToScreenLog(dsx_log, str(patrons_current) + ' patrons are current')

		printToScreenLog(dsx_log, '')
		# delete old_patron_load if it exists
		printToScreenLog(dsx_log, 'Deleting previous patron load')
		if old_patron_load:
			os.remove(old_patron_load)
		# rename new_patron_load to 'old-'+new_patron_load
		printToScreenLog(dsx_log, 'Saving latest patron load')
		os.rename(new_patron_load, os.getcwd() + '\\old-' + os.path.basename(new_patron_load))
		# copy the completed dsp api file to the dsx shared directory
		printToScreenLog(dsx_log, 'Copying DSX import file to DSX host')
		shutil.copy(dsx_api_filename, dsx_host)
		shutil.copy(dsx_api_filename, dsx_api)
		# delete dsx_api_filename
		printToScreenLog(dsx_log, 'Removing local DSX import file')
		os.remove(dsx_api_filename)
	else:
		printToScreenLog(dsx_log, 'No new patron load found!')

	end_time = timedelta(seconds=time.monotonic() - start_time)
	printToScreenLog(dsx_log, '\nTime taken to run the process: ' + str(timedelta(seconds=time.monotonic() - start_time)))

dsx_log.close()
input("Press Enter to continue...")
