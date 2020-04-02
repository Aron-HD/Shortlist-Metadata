import csv
import os
from shutil import copyfileobj
from pathlib import Path

"""This script successfully merges .htm and .txt files, adding campaign details templates to each ID.
Current problem is that it creates a new .htm file when only the .txt campaign details are present.
missing .htm files are usually due to there being dupes."""

def main():

	dirName = f'merged'

	if not os.path.exists(dirName): # change to pathlib
	   os.makedirs(dirName)    

	with open ('raw_metadata.csv', 'r') as df:
		csv_data = csv.DictReader(df)

		num = 0
		
		for line in csv_data:
			ID = line['ID']
			head, tail = ID.split('.')
			ID = head
			firstfile = Path(rf'C:\Users\arondavidson\Desktop\Testing Scripts\Scripts\Metadata\txt\\{ID}.txt') 
			secondfile = Path(rf'C:\Users\arondavidson\Desktop\Testing Scripts\Scripts\Metadata\htm\\{ID}.htm')

			fn = fr'{dirName}\\{ID}.htm'
			with open(fn, "wb") as wfd:
				for f in [firstfile, secondfile]:
					try: 
						with open(f, "rb") as fd:
							copyfileobj(fd, wfd, 1024 * 1024 * 10) 
							num += 0.5
					except FileNotFoundError:
						print(f"\tx '{ID}.htm' not found - x (likely a dupe)")
						num -= 0.5


	print(f'\n- {num} files created...')

if __name__ == '__main__':
	main()