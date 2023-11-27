#!/usr/bin/env python3

#### This code requires pandas

import os
import sys
import pandas as pd
import fnmatch
import shutil
from shutil import copytree
import warnings
warnings.simplefilter("ignore")

#### Test for working directory

path = str(input('\nAre your input folders & metadata in the same working directory? Y/N\n')).upper()
if path == 'Y':
	cwd = os.getcwd()
else:
	print('Move contents to working current directory and try again')
	sys.exit(1)

#### Data tidying

try:
	raw_data = str(input('\nEnter excel file containing metadata\nFormat: filename.xlsx\n'))
	dataframe = pd.read_excel(raw_data)
	write_dataframe = dataframe[['Complaint Number','Requester Company']].to_string(index=False)
	outfile = open('metadata.txt','w')
	outfile.write(write_dataframe)
	outfile.close()
except IOError as er:
    print('Invalid Entry', str(er))
    sys.exit(1)

#### Registering information

infile = open('metadata.txt','r')
headline = infile.readline()

complaint_dir = {}
for line in infile:
	line = line.rsplit()
	key = line[0]
	if line[1] == 'NaN':
		value = 'Region Unregistered'
		complaint_dir[key]=value
	else:
		value =line[1]
		complaint_dir[key]=value
infile.close()

#### Get folders containing techline information

tech_dir = set()
for root, dirs, files in os.walk(cwd):
	for filename in fnmatch.filter(files,'insert_pattern*.zip'):
 		tech_dir.add(root.split('/')[-1])

#### Check if folders and the metadata correspond

match_list = []
unmatch_list = []
for element in tech_dir:
	if element in complaint_dir:
		match_list.append(element)
	else:
		unmatch_list.append(element)
	if match_list == []:
		print('\nThe folder names and metadata do not match, please check and rerun.')
		sys.exit(1)
print('\nThese folders (for files:insert_pattern) I found are listed in the metadata:\n',match_list,'\nThese folders (for files:insert_pattern) are not listed in metadata:\n',unmatch_list)

#### Create main directory

print('\n........Creating folder named -Sorted-')
nwd = cwd+'/Sorted'
if os.path.exists(nwd):
	ask = str(input('\nFolder exists, overwrite?: Y/N\n')).upper()
	if ask == 'Y':
		shutil.rmtree(nwd)
	else:
		print('\nEnding Program')
		sys.exit(1)
		
os.mkdir(nwd)
os.chdir(nwd)
for j in tech_dir:	
	if os.path.isdir(complaint_dir[j]):
		dupe_cwd = nwd + '/' + complaint_dir[j]
		os.chdir(dupe_cwd)
		os.mkdir(j)
		source = cwd+'/'+j
		destination = dupe_cwd + '/' + j
		copytree(source,destination, dirs_exist_ok = True)
	else:
		os.mkdir(complaint_dir[j])
		snwd = nwd+'/'+complaint_dir[j]
		os.chdir(complaint_dir[j])
		os.mkdir(j)
		source = cwd+'/'+j
		os.chdir(j)
		destination = os.getcwd() 
		copytree(source,destination, dirs_exist_ok = True)
	
	snwd = ''
	os.chdir(nwd)