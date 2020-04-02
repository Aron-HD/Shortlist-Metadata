import pandas as pd
from glob import iglob
from pathlib import Path


def main():
	'''
	grabs the sheets (except first sheet) from 
	'shortlist metadata.xslx' excel sheet in 
	the	folder stated through input.
	'''
	
	# user input path to metadata xlsx file
	path = r'T:\Ascential Events\WARC\Backup Server\Loading\Monthly content for Newgen\Project content - March 2020\WARC Awards 2020'#input('paste folder path: ') # logger.info('user input: ' + path)

	# filter only xlsx files matching shortlist metadata in the filename
	files = iglob(fr'{path}\\*hortlist*metadata*.xlsx')

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
					print(sheet)
					# make csv files in correct folder
					csvf = Path(f'{newdir}\\{sheet}_metadata.csv')
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

if __name__ == '__main__':
	main()