#!/usr/bin/env python
# -*- coding: ASCII -*-

"""mybackupfiles.py: Script to copy files and compress them and put them in a separate location. """

__version__ = "0.0.2"

__author__= "Christopher Weller"
__copyright__="Copyright 2009, Planet Earth"
__credits__= ["Christopher Weller"]
__license__= "GPL"
__contact__="Christopher Weller"
__deprecated__="False"
__date__="13 April 2017"
__maintainer__= "Christopher Weller"
__email__= "christopher.weller@gm.com"
__status__= "Production"


# Written in Python 2.7

import os
import zipfile
import shutil

#copy files and folder and compress into a zip file
def	doprocess(source_folder, target_zip):
	zipf = zipfile.ZipFile(target_zip, "w")
	for subdir, dirs, files in os.walk(source_folder):
		for file in files:
			print (os.path.join(subdir, file))
			zipf.write(os.path.join(subdir, file))
	
	print ("Created ", target_zip)
	

#copy files to a target folder	
def	docopy(source_folder, target_folder):
	for subdir, dirs, files in os.walk(source_folder):
		for file in files:
			print (os.path.join(subdir, file))
			shutil.copy2(os.path.join(subdir, file), target_folder)
	
		

if __name__ =='__main__':
	print ('Starting execution')
	
	#compress to zip
	source_folder = 'c:\\Users\\qzf681\\Documents'
	target_zip = 'e:\\backups\\BackupData.zip'
	doprocess(source_folder, target_zip)	
			
	#copy to backup folder
	source_folder = 'c:\\Users\\qzf681\\Documents'
	target_folder = 'e:\\backups\\DailyBackup'
	docopy(source_folder, target_folder)
	
	
	print ('Ending execution')
