import os 
from docx2pdf import convert
from pdf2image import convert_from_path
import glob
import win32com.client
doc_name = ''
choise0 = input('Convert to docx? |: ')
if choise0 == 'y':
	word = win32com.client.Dispatch("Word.Application")
	for i, doc in enumerate(glob.iglob("*.doc")):
		in_file = os.path.abspath(doc)
		wb = word.Documents.Open(in_file)
		out_file = os.path.abspath(f'{doc[:-4]}.docx'.format(i))
		wb.SaveAs2(out_file, FileFormat=16) # file format for docx
		wb.Close()
	word.Quit()	
print(doc_name)
#-----------------------------------------------------------------------------------------------
choise = input('Convert to pdf? |: ')
if choise == 'y':
	convert('.') 
#-----------------------------------------------------------------------------------------------
path = '.'
files = os.listdir(path)
png_counter = 1
pdf_files = []
for index in files:
	if(index[-3:] == 'pdf'):
		pdf_files.append(index)
doc_files = []
for index in files:
	if(index[-3:] == 'doc'):
		doc_files.append(index)
choise2 = input('Convert to png? |: ')
if choise2 == 'y':
	for index in pdf_files:
		images = convert_from_path(index, 700)
		for image in images:
			image.save(f'{index[:-4]}_{png_counter}.png')
			png_counter+=1
		png_counter=1