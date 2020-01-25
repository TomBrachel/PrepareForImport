# # -----for all vol product - Fix Columns

import pandas as pd
import xlrd
import os, sys
from shutil import copyfile
from pyxlsb import open_workbook as open_xlsb

# ===CONSTANTS:
vol_TV_names = ['AVBOX', 'AVMBX', 'AVONM', 'AVHDA', 'AVEMI', 'AVHEI', 'AVHEL', 'AVOEM', 'AVUNI', 'VOUNI']
vol_app_names = [ 'VMBOX', 'VMBXS', 'VONMC', 'VHDAR', 'VEMIA', 'VHELI', 'VHELU', 'VOEMI']
ft_names = ['ARBFD', 'BOXFD', 'BOXL2', 'BXNA2', 'MBOX2', 'MBXSO', 'NMCFT', 'HEDSM', 'HS100', 'EMIFT', 'ERBFT', 'HELFT', 'HUNFD', 'UNIFT', 'NMCFU', 'EA1FT', 'EDEN2', 'TAKT2', 'TPCOF', 'UNCL2']
ft_stack_names = ['ARBFDN', 'BOXFDN', 'BOXL2N', 'BXNA2N', 'MBOX2N', 'MBXSON', 'NMCFTN', 'HEDSMN', 'HS100N', 'EMIFTN', 'ERBFTN', 'HELFTN', 'HUNFDN', 'UNIFTN', 'NMCFUN', 'EA1FTN', 'EDEN2N', 'TAKT2N', 'TPCOFN', 'UNCL2N']
unicell_names = ['EA1FT', 'EDEN2', 'TAKT2', 'TPCOF', 'UNCL2', 'EA1FTN', 'EDEN2N', 'TAKT2N', 'TPCOFN', 'UNCL2N']
origin_path = os.getcwd()
owner_by_number_dict = { '682424' : 'mBox', 
						 '677208' : 'NMC' ,
						 '678522' : 'HedArtzi',
						 '682371' : 'Helicon' ,
						 '678101' : 'Unicell' }
#==========================
df_cellcom_verafication = pd.DataFrame(columns=['Owner', 'Sheet', 'Dwonloads', 'Charge', 'Revenu', '%'])
#takes sheet_name and str
#return String for new name to sheet with the str in right place
def addToNameSheet(sheet_name, str):
	splitted_sheet_name = sheet_name.split('_')
	splitted_sheet_name[1] += str
	return '_'.join(splitted_sheet_name)

def isContains(a, b):
	try:
		if a in b:
			return True
	except:
		return False

	return False

#get sheet name 
#return dict of headers
def get_origin_cols(sheet_name):
	if isVolApp(sheet_name) or isVolTV(sheet_name):
		return {'Unnamed: 0': 'Item ID ',
					  'Unnamed: 1': ' Application Code',
		 			  'Unnamed: 2': ' Customer Type',
		 			  'Unnamed: 3': ' Item Description', 
		 			  'Unnamed: 4': ' Item Description Hebrew',
		 			  'Unnamed: 5': ' Performer',
		 			  'Unnamed: 6': ' ISRC Code',
		 			  'Unnamed: 7': ' ACUM Code', 
		 			  'Unnamed: 8': ' Num of Events',
		 			  }
	else:
		return {'Unnamed: 0': 'Item ID ',
					  'Unnamed: 1': ' Application Code',
		 			  'Unnamed: 2': ' SM Type', 
		 			  'Unnamed: 3': ' Customer Type',
		 			  'Unnamed: 4': ' Item Description', 
		 			  'Unnamed: 5': ' Item Description Hebrew',
		 			  'Unnamed: 6': ' Performer',
		 			  'Unnamed: 7': ' ISRC Code',
		 			  'Unnamed: 8': ' ACUM Code', 
		 			  'Unnamed: 9': ' Num of Events',
		 			  'Unnamed: 10': ' Charge' }

#takes sheet_name
#return True if is VolApp sheet
def isUnicellFT (sheet_name):
	for s in unicell_names:
		if s in sheet_name:
			return True
	return False


#takes sheet_name
#return True if is VolApp sheet
def isFT (sheet_name):
	for s in ft_names:
		if s in sheet_name:
			return True
	return False

#takes sheet_name
#return True if is stack name sheet
def isFTStack (sheet_name):
	for s in ft_stack_names:
		if s in sheet_name:
			return True
	return False



#takes sheet_name
#returns True if is VolApp sheet
def isVolApp (sheet_name):
	for vol in vol_app_names:
		if vol in sheet_name:
			return True
	return False

#takes sheet_name 
# returns True if is VolTV sheet
def isVolTV (sheet_name):
	for vol in vol_TV_names:
		if vol in sheet_name:
			return True
	return False

def getPercent (df_cellcom_percentage, sheet_name):
		try:
			percentage_dict = df_cellcom_percentage.set_index('Sheet').T.to_dict('list')
			sheet = sheet_name.split('_')[1]
			return float(percentage_dict[sheet][0])
		except Exception as e:
			return 1
			print('getPercent failed - ' + str(e))


#takes sheet name
#returns original name (without 'N' or 'PB'\'PL')
def resetName(sheet_name):
	sheet = sheet_name.split('_')[1]

	if isVolApp(sheet_name):
		return sheet[:len(sheet)-2]
	elif isVolTV(sheet_name):
		return sheet 
	elif isFTStack(sheet_name):
		return sheet[:(len(sheet) - 1)]

	return ''

#takes the html summery and sheet name
#returns charge for this sheet
def getTotal(df_html_list, sheet_name):
	#if helicon:
	if len(df_html_list) == 1:
		start_summery_row = 0
		end_summery_row = 0
		df_summery = df_html_list[0]
		for i, k in df_summery.iterrows():
			if df_summery.iat[i, 0] == 'Carrier' and start_summery_row == 0:
				start_summery_row = i
			elif df_summery.iat[i, 0] == 'Total Revenue Sharing' and end_summery_row == 0:
				end_summery_row = i

		df_summery = df_summery.iloc[start_summery_row:end_summery_row, [0,1,2]].dropna()
		summery_dict = df_summery.set_index(0).T.to_dict()
		sheet = resetName(sheet_name)
		agrr = summery_dict[sheet][1]

		df_summery = df_html_list[0]
		for i, k in df_summery.iterrows():
			if df_summery.iat[i, 0] == 'PARTNER:' and isContains(agrr, df_summery.iat[i, 1]):
				j = 1
				while df_summery.iat[i + j, 0] != 'Total Revenue Sharing':
					j += 1
				total = float(df_summery.iat[i + j - 1, 4])
				if df_summery.iat[i + j - 1, 0] == '20perc deduction':
					total += float(df_summery.iat[i + j - 2, 4])
				if isVolApp(sheet_name):
					return total / 2
				return total

		return 0
	#if not helicon:
	else:
		try:
			df_summery = df_html_list[1]
			df_summery = df_summery[[0,1,2]].dropna()
			summery_dict = df_summery.set_index(0).T.to_dict()
			sheet = resetName(sheet_name)
			agrr = summery_dict[sheet][1]

			if isVolApp(sheet) or isVolTV(sheet):
				return float(summery_dict[sheet][2])
			for i, df in enumerate(df_html_list[2:]):
				if str(df.iat[0,0]).startswith('PARTNER: ' + agrr + ' START DATE:', 0) and ('DUMMY' not in df.iat[0,0]):
					curr_df = df_html_list[i + 4][[3,4]]
					for j, k in curr_df.iterrows():
						if curr_df.at[j, 3] == 'Total Revenue Sharing':
							charge = float(curr_df.at[j - 1,4])
							if isVolApp(sheet_name):
								return charge / 2
							return charge
			return 0

		except Exception as e:
			print('getTotal failed - ' + str(e))
			return 0
	



def createFileList(ext_file):
	file_list = []
	for file in os.listdir(os.getcwd()):
    		if (file.endswith(ext_file)):
    			current_file = str(os.path.join(os.getcwd(), file))
    			file_list.append(current_file.split('/')[-1])
	return file_list

def getIdFromFile(file_name):
	if file_name.endswith('.xls'):
		return file_name.split('_')[2]
	elif file_name.endswith('.html'):
		return file_name.split('_')[1]
	if file_name.endswith('.xlsb'):
		if file_name.startswith('DT', 0):
			return file_name.split('_')[2]
		elif file_name.startswith('DT', 0):
			return file_name.split('_')[1]

def createDictFiles (xls_list, html_list):
	file_dict = {}
	for xls in xls_list:
		xls_id = getIdFromFile(xls)
		for html in html_list:
			if getIdFromFile(html) == xls_id:
				file_dict.update( { xls : html } )
	return file_dict

# fix cols and split sheets
def fixColAndSplitSheets(xls_file):
	print('=======fix cols and split sheets======')
	with pd.ExcelWriter(xls_file) as writer:
		rb = xlrd.open_workbook(xls_file)
		for s in rb.sheets():
			df = pd.read_excel(xls_file, sheet_name=s.name) #Read Excel file as a DataFrame
			if owner_by_number_dict[getIdFromFile(xls_file)] == 'Helicon':
				df1 = pd.DataFrame([[''] * len(df.columns)], columns=df.columns)
				df = df1.append(df, ignore_index=True)
			# fix headers for vol sheets
			if (isVolApp(s.name) or isVolTV(s.name)):
				df.insert(2,' SM Type', '') 
				df.at[4,' SM Type'] = ' SM Type'
				df.insert(10,' Charge', '') 
				df.at[4,' Charge'] = ' Charge'
				df.iat[4, 8] = ' ACUM Code'
				for i, trial in df.iterrows():
					if df.iat[i, 1] == 'Music':
						df.iat[i, 1] = 'Playback'	


			#fix headers of df
			originCols = get_origin_cols(s.name)
			df.rename(columns=originCols, inplace=True)

			# delete sum rows
			without_sum_rows = []
			for cust_type in df[' Customer Type']:
				try:
					str_cust_type = str(cust_type)
					if 'Sum Of:' in cust_type:
						without_sum_rows.append(False)
					else:
						without_sum_rows.append(True)
				except:
					without_sum_rows.append(True)
			df = df[without_sum_rows]

			#split volApp sheets to Playlist and Playback - and append them as new sheets.
			if isVolApp(s.name):
				df_playback = df[df[' Application Code'] != 'Playlist']
				df_playlist = df[df[' Application Code'] != 'Playback']

				df_playback.to_excel(writer, sheet_name=addToNameSheet(s.name, 'PB'), index=False)#, header=False)
				df_playlist.to_excel(writer, sheet_name=addToNameSheet(s.name, 'PL'), index=False)#, header=False)
			
			elif isFT(s.name):
				df_FunDial = df[df[' Application Code'] != 'FD_NEW']
				df_FD_NEW = df
				if isUnicellFT(s.name):
					df_FunDial = df_FunDial[df_FunDial[' Customer Type'] != 'Cellcom Employee']
					df_FD_NEW = df_FD_NEW[df_FD_NEW[' Application Code'] != 'FunDial']
					df_FD_NEW = df_FD_NEW[df_FD_NEW[' Customer Type'] != 'Cellcom Employee']

				df_FunDial.to_excel(writer, sheet_name=s.name, index=False)
				df_FD_NEW.to_excel(writer, sheet_name=addToNameSheet(s.name, 'N'), index=False)#, header=False)
			else:
				df.to_excel(writer, sheet_name=s.name, index=False)#, header=False)
			print(s.name)



def fixSumAndPlays(xls_file, html_file):
	
	print('=======fix sum and plays for every sheet======')
	# sum
	# getting the .html file
	print(html_file)
	try:
		df_html_list = pd.read_html(html_file)
	except Exception as e:
		print('Can\'t find html file ' + str(e))
		exit(1)


	#getting the percentage .csv
	try:
		df_cellcom_percentage = pd.read_csv(origin_path + '/cellcom_percentage.csv')
	except Exception as e:
		print('Can\'t find cellcom_percentage.csv - ' + str(e))
		exit(1)

	with pd.ExcelWriter(xls_file) as writer:
		rb = xlrd.open_workbook(xls_file)
		for s in rb.sheets():
			df = pd.read_excel(xls_file, sheet_name=s.name)
			print(s.name)
			# if isFT(s.name) or isVolApp(s.name) or isVolTV(s.name):
			sum_plays = 0
			sum_charge = 0
			sum_charge_re = 0
			percent = getPercent(df_cellcom_percentage, s.name)
			for i, trial in df.iterrows():

				if i > 4 and df.at[i, 'Item ID '] != 'Report Summery':
					if isFTStack(s.name):
						df.at[i, ' Application Code'] = 'FD_NEW'



					try:
						sum_plays += df.at[i, ' Num of Events']
						sum_charge += df.at[i, ' Charge']
					except Exception as e:
						peint('failed to get plays and charge - ' + str(e))
						sum_charge += 0
				elif df.at[i, 'Item ID '] == 'Report Summery':
					df.at[i, ' Num of Events'] = sum_plays
					sum_charge_re = sum_charge * percent
					if isFTStack(s.name) or isVolTV(s.name) or isVolApp(s.name):
						sum_charge = getTotal(df_html_list, s.name)

						sum_charge_re = sum_charge
						sum_charge = sum_charge / percent


					df.at[i, ' Charge'] = sum_charge

			if isFTStack(s.name) or isVolTV(s.name) or isVolApp(s.name):
				charge_per_play = sum_charge / sum_plays
				for i, trial in df.iterrows():
					if i > 4 and df.at[i, 'Item ID '] != 'Report Summery':
						df.at[i, ' Charge'] = df.at[i, ' Num of Events'] * charge_per_play
			if sum_plays > 0:
				sheet = s.name.split('_')[1]
				df_row = pd.DataFrame([[owner_by_number_dict[getIdFromFile(xls_file)],
				 						sheet,
				 						sum_plays, 
				 						sum_charge, 
				 						sum_charge_re, 
				 						percent]], columns=['Owner', 'Sheet', 'Dwonloads', 'Charge', 'Revenu', '%'])
				global df_cellcom_verafication
				df_cellcom_verafication = df_cellcom_verafication.append(df_row, ignore_index=True)
				


			df.to_excel(writer, sheet_name=s.name, index=False, header=False)

	print(xls_file + ' is ready for import!')

def xlsbToXls(xlsb_file):
	new_file_name = xlsb_file.split('.xlsb')[0] + '.xls'
	df = []
	with pd.ExcelWriter(new_file_name) as writer:
		with open_xlsb(xlsb_file) as wb:
			for sheet in wb.sheets:
				df = []
				for row in wb.get_sheet(sheet).rows():
					df.append([item.v for item in row])
				df = pd.DataFrame(df[1:])
				df.to_excel(writer, sheet_name=sheet, index=False, header=False)
	return new_file_name

def xlsbToHtml(xlsb_file):
	new_file_name = xlsb_file.split('.xlsb')[0] + '.html'
	df = []
	# with pd.ExcelWriter(new_file_name) as writer:
	with open_xlsb(xlsb_file) as wb:
		for sheet in wb.sheets:
			df = []
			for row in wb.get_sheet(sheet).rows():
				df.append([item.v for item in row])
			df = pd.DataFrame(df[1:])
	df.dropna()
	df.to_html(new_file_name, index=False, header=False)
	return new_file_name

def getDateFromFileName(file_name):
	date = file_name.split('_')[-1]
	year = date[0:4]
	month = date[4:6]
	print(date)
	print(year)
	print(month)
	return month + "." + year

#START===========================================================================================START

# mapping all relevent files
print('mapping all relevent files..')
xls_list = createFileList('.xls')
html_list = createFileList('.html')
xlsb_list = createFileList('.xlsb')
xlsb_DT = ''
xlsb_ic = ''

#mapping xlsb files to DT and ic
for xlsb in xlsb_list:
	if xlsb.startswith('DT', 0):
		xlsb_DT = xlsbToXls(xlsb)
	elif xlsb.startswith('ic', 0):
		xlsb_ic = xlsbToHtml(xlsb)

#make new dir if doesnt exsist
if not os.path.exists(origin_path + '/readyToImport'):
        os.makedirs(origin_path + '/readyToImport')

#copy all xls to new folder to work on them
print('copy files to /readyToImport..')
for xls in xls_list:
	src_xls = os.getcwd() + '/' + xls
	dest_xls = origin_path + '/readyToImport' + '/' + xls
	copyfile(src_xls, dest_xls)


file_dict = createDictFiles(xls_list, html_list)


if xlsb_ic and xlsb_DT:

	src_xlsb = os.getcwd() + '/' + xlsb_DT
	dest_xlsb = origin_path + '/readyToImport/' + xlsb_DT
	copyfile(src_xlsb, dest_xlsb)
	file_dict.update({xlsb_DT : xlsb_ic})


os.chdir(origin_path + '/readyToImport')
for xls in file_dict:

	print(owner_by_number_dict[getIdFromFile(xls)] + ' File:')
	fixColAndSplitSheets(xls)
	html = origin_path + '/' + file_dict[xls]
	fixSumAndPlays(xls, html)
	print()



date = getDateFromFileName(list(file_dict.keys())[0])
if not os.path.exists('Reports Verification_' + date + '.xlsx'):
	print("creating Reports Verification file..")
	df_cellcom_verafication.to_excel('Reports Verification_' + date + '.xlsx', sheet_name='Cellcom', index=False)
	print("Reports Verification file updated")
else:
	with pd.ExcelWriter('Reports Verification_' + date + '.xlsx') as writer:
		

		df_report = pd.read_excel('Reports Verification_' + date + '.xlsx', sheet_name=None)
		df_report['Cellcom'] = df_report['Cellcom'].append(df_cellcom_verafication)
		for sheet in df_report:
			df_report[sheet].to_excel(writer, sheet_name=sheet, index=False)
	print("Reports Verification file updated")





