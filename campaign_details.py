import csv
import os # from pathlib import Path
from excel_to_csv import output_file

dirName = f'txt'

def main():
	"""
	This script writes wraps the campaign details template around each row of the 'raw_metadata.csv' and writes it into a separate text file named after each ID in the 'htm_files' folder.

	Currently I can combine these with the html files that we get back from newgen so that we don't have to copy and paste the template.

	I want to be able to write this directly into the file though, so that we can overwrite the template if it is there. 

	Though this won't be necessary if newgen are asked not to include the template.
	"""	

	if not os.path.exists(dirName):
	    os.makedirs(dirName)

	number = 0    

	with open (output_file, 'r') as df:
		csv_data = csv.DictReader(df)
		
		for line in csv_data:
			ID_ext = line['ID']
			ID = ID_ext.split('.')[0]
			Brand = line['Brand']
			Owner = line['Brand Owner Name']
			Lead_agency = line['Lead Agencies']
			Contributing_agency = line['Contributing Agencies']
			Market = line['Countries']
			Industries = line['Industry Sectors']
			Media_channels = line['Media channels']
			Budget = line['Budget']

			fn = fr'{dirName}\\{ID}.txt'

			with open (fn, 'w') as f:
			
				template = f"""\
<html>
<body>
<!--StartofArticle {ID}-->
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
				
				print(f"- created '{ID}.txt'")

				number += 1

	print(f"""
- {number} sets of campaign details created in '{dirName}'

##############################################################

#~~~~~~~~~~~~~~~~~~~~~ SCRIPT FINISHED ~~~~~~~~~~~~~~~~~~~~~~#

##############################################################

#	IMPROVEMENTS
#	
#	- use pandas df.loc['Column name'] to read this directly from Sian's 
#	  metadata xlsx, removing the need for the 'excel_to_csv.py' module
#	  (although, it might be more difficult to do individual chunks of files.)
#
#	- add logging and try, except to catch errors.
#
""")

if __name__ == '__main__':
	files_in_dirName = len(os.listdir(dirName))
	print(f"\n#			Pre-script: {files_in_dirName} files in 'txt_files' folder")

	main()

	files_in_dirName = len(os.listdir(dirName))
	print(f"\n#			Post-script: {files_in_dirName} files in 'txt_files' folder\n")
	