import csv
import pandas as pd
from pathlib import Path
from glob import glob
from get_sheets import path

dirName = f'txt'
	
def metadata(csvfn, dupes):
	"""
	This script writes wraps the campaign details template around each row of the '{category}_metadata.csv' and writes it into a separate text file named after each ID in the 'htm_files' folder.

	Currently I can combine these with the html files that we get back from newgen so that we don't have to copy and paste the template.

	I want to be able to write this directly into the file though, so that we can overwrite the template if it is there. 

	Though this won't be necessary if newgen are asked not to include the template.
	"""
	
	if not Path(dirName).is_dir(): 
		dirName.mkdir() # logger.info("")

	number = 0    

	for file in glob(f'csv\\*{csvfn}*.csv'):
		print('\nreading: ' + file)

		with open(Path(file)) as df:
			csv_data = csv.DictReader(df)
			
			for line in csv_data:
				
				ID = line['ID']
				Brand = line['Brand']
				Owner = line['Brand Owner Name']
				Lead_agency = line['Lead agencies']
				Contributing_agency = line['Contributing agencies']
				Market = line['Countries']
				Industries = line['Industry sector']
				Media_channels = line['Media used']
				Budget = line['Budget']

				# write a txt file with details if not a dupe
				if not int(ID) in dupes:
					fn = fr'{dirName}\\{ID}.txt'		
					with open (fn, 'w', encoding='1252') as f:
						template = f"""\
<html>
<body>
<!-- {ID} Campaign Details -->
<h3>Campaign details</h3>
<p><strong>Brand:</strong> {Brand}<br/>
<strong>Brand owner:</strong> {Owner}<br/>
<strong>"""

						if not (len(Contributing_agency) > 0):
							if ',' in Lead_agency:
								template += f"Agencies:</strong> {Lead_agency}<br/>"
							else:
								template += f"Agency:</strong> {Lead_agency}<br/>"

						if (len(Contributing_agency) > 0):
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

def pre_post_script():
	files_in_dirName = len(glob(dirName + '\\*.txt'))
	return f"- {files_in_dirName} txt files in 'txt_files' folder"

def main():
	
	print('			# PRE-SCRIPT: ' + pre_post_script())

	# find 'Dupes' tab in main EDIT spreadsheet and save all IDs in tab to a list to check against
	df = pd.read_excel(f'{path}\\WARC Awards_EDIT.xlsx', sheet_name='Dupes')
	dupes = df['ID'].tolist()

	# categories to loop through (change to pandas and read tabs from shortlist metadata.xlsx instead of csv files)
	for csvfn in ['Content', 'Social']: #  'Innovation', 'Purpose'
		metadata(csvfn, dupes)

	print(f"""

##############################################################

#~~~~~~~~~~~~~~~~~~~~~ SCRIPT FINISHED ~~~~~~~~~~~~~~~~~~~~~~#

##############################################################

#	IMPROVEMENTS
#	
#	- use pandas df['Column name'] to read this directly from Sian's 
#	  metadata xlsx, removing the need for the 'get_sheets.py' module
#	  (although, it might be more difficult to do individual chunks of files.)
#
#	- add logging and try, except to catch errors.
#

###############################################################
""")

	print('			# POST-SCRIPT: ' + pre_post_script())

if __name__ == '__main__':
	main()

	
	