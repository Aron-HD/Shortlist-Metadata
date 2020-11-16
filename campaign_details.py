import pandas as pd, PySimpleGUI as sg
from shutil import copyfileobj
from pathlib import Path
from glob import glob

def campaign_details(txtdir, dupes, excel_file, sheets):
	
	if not Path(txtdir).is_dir(): 
		txtdir.mkdir() # logger.info("")

	number = 0
	
	for file in excel_file:
		for sheet in sheets:
			df = pd.read_excel(file, sheet_name=sheet)
			for index, row in df.iterrows():
				ID = row['ID']
				Brand = row['Brand']
				Owner = row['Brand owner']
				Lead_agency = row['Lead agencies']
				Contributing_agency = str(row['Contributing agencies'])
				Market = row['Market']
				Industries = row['Industry sector']
				Media_channels = row['Media channels']
				Budget = row['Budget']
				# write a txt file with details if not a dupe
				if not ID in dupes:
					fn = fr'{txtdir}/{ID}.txt'		
					with open (fn, 'w', encoding='1252') as f:
						template = f"""\
<html>
<body>
<!-- {ID} Campaign Details -->
<h3>Campaign details</h3>
<p><strong>Brand:</strong> {Brand}<br/>
<strong>Brand owner:</strong> {Owner}<br/>
<strong>"""

						if Contributing_agency == 'nan':
							if ',' in Lead_agency:
								template += f"Agencies:</strong> {Lead_agency}<br/>"
							else:
								template += f"Agency:</strong> {Lead_agency}<br/>"

						if Contributing_agency != 'nan':
							if ',' in Lead_agency:
								template += f"Lead agencies:</strong> {Lead_agency}<br/>"
							if not ',' in Lead_agency:
								template += f"Lead agency:</strong> {Lead_agency}<br/>"
							if ',' in Contributing_agency:
								template += f"\n<strong>Contributing agencies:</strong> {Contributing_agency}<br/>"
							if not ',' in Contributing_agency:
								template += f"\n<strong>Contributing agency:</strong> {Contributing_agency}<br/>"

						template += f"""
<strong>Market:</strong> {Market}<br/>
<strong>Industries:</strong> {Industries}<br/>
<strong>Media channels:</strong> {Media_channels}<br/>"""

						if 'onfidential' in Budget:
							pass
						elif Budget == '0k':
							template += f"\n<strong>Budget:</strong> No budget</p>"
						else:
							template += f"\n<strong>Budget:</strong> {Budget}</p>"

						f.write(template)
						number += 1
						print(f"- created '{ID}.txt'")
				else:
					print(f"- x dupe '{ID}'")
		print(f"\n- {number} sets of campaign details created in '{txtdir}'")

def merge(excel_file, mergedir, dupes, murkies, sheets):
	
	num = 0
	dup = 0

	for file in excel_file:
		for sheet in sheets:
			df = pd.read_excel(file, sheet_name=sheet)
			IDs = df['ID'].astype(int).tolist()

	for ID in IDs:

		if not ID in dupes and not ID in murkies:

			firstfile = Path(fr'txt/{ID}.txt') 
			secondfile = Path(fr'htm/{ID}.htm')
			# tidy htm
			
			with open(secondfile, 'r') as f:
				contents = f.read()
				try:
					# can't get regex to work
					# search = re.findall(r'<sup>\d*</sup>[\.\^\$\*\+\?\[\]\{\}\(\)\-@#~/"\':;=£%&!]', contents)
					# print(ID, search)
					# for old in search:
					# 	head = old[:-1]
					# 	tail = old[-1]
					# 	new =f'{tail}{head}'
					# 	# print(new)
					out = contents.replace('<html>', '\n').replace('<body>', '')
					with open(secondfile, 'w') as f:
						f.write(out)
					print('tidied ' + str(ID))

				except Exception as e:
					raise e

			fn = Path(mergedir) / fr'{ID}.htm'
			with open(fn, "wb") as wfd:
				for f in [firstfile, secondfile]:
					try: 
						with open(f, "rb") as fd:
							copyfileobj(fd, wfd, 1024 * 1024 * 10)
						
					except FileNotFoundError:
						print(f"\tx '{ID}.htm' not found - x (likely a dupe)")
				print(fn.name)
			num += 1

		else:
			dup += 1

	print(f'\n- {num} files... [{dup} dupes or murkies] ')

def get_sheets(dupes, murkies, txtdir, mergedir, excel_file):

	for f in excel_file:
		xl = pd.ExcelFile(f) # logger.info("")
		sheets = xl.sheet_names

	sg.theme('DarkPurple4')
	layout = [
		[sg.Checkbox(sheet, change_submits=True, key=sheet) for sheet in sheets],
		[sg.Button('Details'), sg.Button('Merge'), sg.Cancel()]
	]
	
	window = sg.Window('Metadata',layout,
    	keep_on_top=True
    )
	while True:
		event, values = window.read()
		if event in ('Cancel', None):
			break
		if event == 'Details':
			campaign_details(dupes=dupes,
							txtdir=txtdir,
							excel_file=excel_file,
							sheets=[v for v in values if values[v] == True]
							)
		if event == 'Merge':
			try:
				print('\nmerging shortlist campaign details with htm')
				merge(dupes=dupes,
					murkies=murkies,
					mergedir=mergedir,
					excel_file=excel_file,
					sheets=[v for v in values if values[v] == True]
					)
			except Exception as e:
				raise e
	window.close()

def get_dupes(path):
	'''
	Get dupes from respective tab in main EDIT spreadsheet
	'''
	files = Path(path).glob(r'*EDIT.xlsx')
	for f in files:
		try:
			df = pd.read_excel(f, sheet_name='Dupes')
			dupes = df['ID'].tolist()
			print()
			print(len(dupes), 'Dupes:', dupes)
		except Exception as e:
			dupes = []
			print('\n', e)
	return dupes

def get_murkies(path):
	'''
	Get murkies from respective tab in main EDIT spreadsheet
	'''
	files = Path(path).glob(r'*EDIT.xlsx')
	for f in files:
		try:
			df = pd.read_excel(f, sheet_name='Murkies')
			murkies = df['ID'].tolist()
			print()
			print(len(murkies), 'Murkies:', murkies)
		except Exception as e:
			murkies = []
			print('\n', e)
	return murkies

def pre_post_script(txtdir, mergedir):
	files_in_txtdir = len(glob(txtdir + r'/*.txt'))
	files_in_mergedir = len(glob(mergedir + r'/*.htm'))

	return f"""
	- {files_in_txtdir} txt files in 'txt' folder
	- {files_in_mergedir} htm files in 'merged' folder
	"""

def main():

	txtdir = 'txt'
	mergedir = 'merged'
	### Select Shortlist or Entrants ###
	sg.theme('DarkPurple4')
	layout = [
	[sg.Text('Paste path to files:')],
	[sg.InputText()],
		[sg.Frame(layout=[
		[sg.Radio('Shortlists', 'R', key='R1', default=True),
		sg.Radio('Entrants', 'R', key='R2', default=False)]],
			title='Metadata',
			title_color='white',
			relief=sg.RELIEF_SUNKEN)],
	[sg.Submit(), sg.Cancel()]
	]
	window = sg.Window('Metadata',layout,
    	keep_on_top=True
    )
	while True:
		event, values = window.read()
		if event in ('Cancel', None):
			break
		if event == 'Submit':
			if len(values[0]) > 0 and '\\' in values[0]:
				try:
					path = values[0]
					dupes = get_dupes(path)
					murkies = get_murkies(path)
				
					print('\n		# PRE-SCRIPT: ' + pre_post_script(txtdir, mergedir))
					
					if values['R1'] == True: # get('R1')
						excel_file = glob(rf'{path}/*hortlist*metadata*.xlsx')
					elif values['R2'] == True:
						excel_file = glob(rf'{path}/*ntrant*metadata*.xlsx')
					# find 'Dupes' tab in main EDIT spreadsheet and save all IDs in tab to a list to check against
					
					get_sheets(dupes=dupes, 
							txtdir=txtdir,
							murkies=murkies, 
							mergedir=mergedir, 
							excel_file=excel_file)

					print('		# POST-SCRIPT: ' + pre_post_script(txtdir, mergedir))

				except Exception as e:
					print("\ncheck path to 'EDIT' and 'metadata' spreadsheets is correct\n", e)
			else:
				print('\ninvalid path entered')


			
	window.close()
	print(f"""

##############################################################

#~~~~~~~~~~~~~~~~~~~~~ SCRIPT FINISHED ~~~~~~~~~~~~~~~~~~~~~~#

##############################################################

#	IMPROVEMENTS
#
#	- Make tickboxes appear in original window and add output 
#	  box, rather than opening a new window.
#

##############################################################
""")
	

if __name__ == '__main__':
	main()

	
	