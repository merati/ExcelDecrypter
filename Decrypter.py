#!/usr/bin/python

import zipfile
import os
import sys, getopt
import shutil


def unzipfile(filename):
	if os.path.isfile(filename) == True:
		if "xlsx" in filename.lower():
			ziploc = "./report/"
		if zipfile.is_zipfile(filename) == True:
			with zipfile.ZipFile(filename, "r") as z:
				z.extractall(ziploc)
		else:
			logging.warning("\n\t[!] Unable to unzip " + str(filename) + ". Possible zip header issue.\n\t[!] If this is between VMs please try copying the file again.")
			sys.exit(2)
	else:
		logging.warning("\n[!] Sorry, cannot locate " + str(filename) + ". Exiting.")
		sys.exit(2)


def zipdir(path, ziph):
	for root, dirs, files in os.walk(path):
		for file in files:
			ziph.write(os.path.join(root, file))
			#print os.path.join(root, file)
def save_file():
	zipf = zipfile.ZipFile("Decrypted-output.xlsx", 'w', zipfile.ZIP_DEFLATED)
	origpath = os.getcwd()
	os.chdir(os.sep.join([origpath, "report"]))
	zipdir('.', zipf)
	zipf.close()
	os.chdir('..')

def remove_protection():
	body = ""
	document = open('./report/xl/workbook.xml', 'r')
	for line in document:
		word = line.split(">")

	for i in range(len(word)-1):
		if not word[i].startswith("<workbookProtection"):
			body += word[i]+">"
	document.close()
	document = open('./report/xl/workbook.xml', 'w')
	document.write(body)
	document.close()

	path = './report/xl/worksheets/'
	dirs = os.listdir(path)
	for file in dirs:
		body = ""
		word = ""
		if file.startswith("sheet"):
			document = open(path+file, 'r')
			for line in document:
				word = line.split(">")
			for i in range(len(word)-1):
				if not word[i].startswith("<sheetProtection"):
					body += word[i]+">"
			document.close()
			document = open(path+file, 'w')
			document.write(body)
			document.close()
	save_file()
	shutil.rmtree('./report')


def main():
	unzipfile(ifile)
	remove_protection()
	print "Decrypted file has been created."


print "This script will remove the password protection from an XLSX file."
print "Usage: ./office-decryptor.py -i <file.xlsx>"
print "Hamedm\n\n\n"

if (os.path.isfile(sys.argv[2])):
	ifile = sys.argv[2]
	main()
