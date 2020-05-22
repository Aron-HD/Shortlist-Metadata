from glob import glob
from pathlib import Path
import pandas as pd, PySimpleGUI as sg

def shortlists(path):
	'''
	grabs the sheets (except first sheet) from 
	'shortlist metadata.xslx' excel sheet in 
	the	folder stated through input.
	'''
	
	# user input path to metadata xlsx file
	

	# filter only xlsx files matching shortlist metadata in the filename
	files = Path(path).glob(r'*hortlist*metadata*.xlsx')

	# make a new directory for shortlists if it doesn't already exist
	newdir = 'csv'
	if not Path(newdir).is_dir(): # logger.info()
		newdir.mkdir() # logger.info("")
	try:
		for file in files:
			p = Path(file)
			# define just filename
			f = p.parts[-1] # logger.info("")
			print(f'grabbing sheets from {f}..')
			try:
				# read each sheet in xlsx file
				xl = pd.ExcelFile(file) # logger.info("")

				for i in range(1,5):
					sheet = xl.sheet_names[i]
					df = xl.parse(i)
					print('Sheet name:', sheet)
					# make csv files in correct folder
					csvf = Path(f'{newdir}/{sheet}_metadata.csv')
					# write each dataframe to a separate csv file in 'shortlists' folder, removing the index
					df.to_csv(csvf, index=False) # logger.info("")
					print(Path.cwd() / csvf)

			except Exception as e:
				# print file and exception
				print('\tx ' + f)
				print("\n\t- Ensure excel filename contains 'shortlist metadata'")
				print(e)

	except Exception as e:
		print("\n\t- Ensure path is correct")
		print(e) # logger.error(e)

def entrants(path):
	files = Path(path).glob(r'*ntrant*metadata*.xlsx')

	# make a new directory for shortlists if it doesn't already exist
	newdir = 'csv'
	if not Path(newdir).is_dir(): # logger.info()
		newdir.mkdir() # logger.info("")

	for file in files:
		try:
			# define just filename
			f = file.parts[-1] # logger.info("")
			print(f"grabbing sheets from '{f}'..")
			try:
				# read each sheet in xlsx file
				xl = pd.ExcelFile(file) # logger.info("")

				sheet = xl.sheet_names[0]
				df = xl.parse(sheet)
				print('Sheet name:', sheet)
				# make csv files in correct folder
				csvf = Path(f'{newdir}/{sheet}_metadata.csv')
				# write each dataframe to a separate csv file in 'shortlists' folder, removing the index
				df.to_csv(csvf, index=False) # logger.info("")
				print(Path.cwd() / csvf)

			except Exception as e:
				# print file and exception
				print('\tx ' + f)
				print("\n\t- Ensure excel filename contains 'entrant', 'metadata'")
				print(e)

		except Exception as e:
			print("\n\t- Ensure path is correct")
			print(e) # logger.error(e)

def main():
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
	window = sg.Window('Get Sheets',layout,
    	keep_on_top=True
    	)
	while True:
		event, values = window.read()

		if event in ('Cancel', None):
			break

		if event == 'Submit':
			try:
				path = values[0]

				if values['R1'] == True: # get('R1')
					shortlists(path)
				elif values['R2'] == True:
					entrants(path)
			except Exception as e:
				print(e)
	window.close()

if __name__ == '__main__':
	main()