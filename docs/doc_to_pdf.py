import sys
import os
import comtypes.client

global word
global wdFormatPDF
global doc

in_file = os.path.abspath(sys.argv[1])
out_file = os.path.abspath(sys.argv[2])

# get extension of in file
ext = in_file.split('.')[-1]

is_word = (ext == 'doc' or ext == 'docx')
is_ppt = (ext == 'ppt' or ext == 'pptx')
is_excel = (ext == 'xls' or ext == 'xlsx')

if is_word:
    wdFormatPDF = 17
    word = comtypes.client.CreateObject('Word.Application')    
    doc =  word.Documents.Open(in_file)
elif is_ppt:
    wdFormatPDF = 32
    word = comtypes.client.CreateObject('Powerpoint.Application')
    doc = word.Presentations.Open(in_file)
elif is_excel:
    wdFormatPDF = 57
    word = comtypes.client.CreateObject('Excel.Application')
    doc = word.Workbooks.Open(in_file)

doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()
