import pandas as pd, PySimpleGUI as sg
from pathlib import Path
from glob import glob

def campaign_details(dirName, path, dupes, excel_file, sheets):
	
	if not Path(dirName).is_dir(): 
		dirName.mkdir() # logger.info("")

	number = 0
	
	for file in excel_file:
		for sheet in sheets:
			df = pd.read_excel(file, sheet_name=sheet)
			for index, row in df.iterrows():
				ID = row['ID']
				Brand = row['Brand']
				Owner = row['Brand Owner Name']
				Lead_agency = row['Lead agencies']
				Contributing_agency = str(row['Contributing agencies'])
				Market = row['Countries']
				Industries = row['Industry sector']
				Media_channels = row['Media used']
				Budget = row['Budget']
				# write a txt file with details if not a dupe
				if not ID in dupes:
					fn = fr'{dirName}/{ID}.txt'		
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
<strong>Media channels:</strong> {Media_channels}<br/>
<strong>Budget:</strong> {Budget}</p>"""

						f.write(template)
						number += 1
						print(f"- created '{ID}.txt'")
				else:
					print(f"- x dupe '{ID}'")
		print(f"- {number} sets of campaign details created in '{dirName}'")


def get_sheets(dirName, path, dupes, excel_file):

	for f in excel_file:
		xl = pd.ExcelFile(f) # logger.info("")
		sheets = xl.sheet_names

	sg.theme('DarkPurple4')
	layout = [
		[sg.Checkbox(sheet, change_submits=True, key=sheet) for sheet in sheets] +
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
			campaign_details(dirName, path, dupes, excel_file, [v for v in values if values[v] == True]) 
	window.close()

def get_dupes(path):
	files = Path(path).glob(r'*EDIT.xlsx')
	for file in files:
		df = pd.read_excel(file, sheet_name='Dupes')
		dupes = df['ID'].tolist()
		print(len(dupes), 'Dupes:', dupes)
	return dupes

def pre_post_script(dirName):
	files_in_dirName = len(glob(dirName + r'/*.txt'))
	return f"- {files_in_dirName} txt files in 'txt_files' folder"

def main():
	dirName = f'txt'
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
			try:
				print('		# PRE-SCRIPT: ' + pre_post_script(dirName))
				path = values[0]
				if values['R1'] == True: # get('R1')
					excel_file = glob(rf'{path}/*hortlist*metadata*.xlsx')
				elif values['R2'] == True:
					excel_file = glob(rf'{path}/*ntrant*metadata*.xlsx')
				# find 'Dupes' tab in main EDIT spreadsheet and save all IDs in tab to a list to check against
				dupes = get_dupes(path)
				get_sheets(dirName, path, dupes, excel_file)
			except Exception as e:
				print(e)
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
	print('		# POST-SCRIPT: ' + pre_post_script(dirName))

if __name__ == '__main__':
	main()

	
	