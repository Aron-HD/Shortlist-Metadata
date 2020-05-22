import pandas as pd
import csv # convert csv section to pandas
from shutil import copyfileobj
from pathlib import Path
from glob import iglob
from campaign_details import path
# import re

"""This script successfully merges .htm and .txt files, adding campaign details templates to each ID.
Current problem is that it creates a new .htm file when only the .txt campaign details are present.
missing .htm files are usually due to there being dupes."""

dirName = f'merged'

def merge(file, dupes):
	
	num = 0
	dup = 0

	# convert csv section to pandas
	with open (file) as df:
		csv_data = csv.DictReader(df)

		for line in csv_data:
			ID = line['ID']

			if not int(ID) in dupes:
				firstfile = Path(rf'txt/{ID}.txt') 
				secondfile = Path(rf'htm/{ID}.htm')
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
						print(e)

				fn = fr'{dirName}\\{ID}.htm'
				with open(fn, "wb") as wfd:
					for f in [firstfile, secondfile]:
						try: 
							with open(f, "rb") as fd:
								copyfileobj(fd, wfd, 1024 * 1024 * 10)
							
						except FileNotFoundError:
							print(f"\tx '{ID}.htm' not found - x (likely a dupe)")
					print(fn.split('\\')[-1])
				num += 1

			else:
				dup += 1

	print(f'\n- {num} files... [{dup} dupes] ')

def main():

	if not Path(dirName).is_dir(): 
		dirName.mkdir()

	f = fr'{path}/*EDIT.xlsx'
	df = pd.read_excel(f, sheet_name='Dupes')
	dupes = df['ID'].tolist()

	# categories to loop through (change to pandas and read tabs from shortlist metadata.xlsx instead of csv files)
	for csvfn in ['Entrants']: # 'Innovation', 'Purpose'
		try:
			for file in iglob(fr'csv/*{csvfn}*.csv'):
				print(f'\nread: {file}')
				merge(file, dupes)
				
		except Exception as e:
			print(e)

	print(f'\n### END ###\n')

if __name__ == '__main__':
	main()
	# tidy htm
