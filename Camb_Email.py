import yagmail
import openpyxl

'''import datetime
x = datetime.datetime.now()
log = x.strftime('%d')+x.strftime('%m')+x.strftime('%Y')+'.txt'
open(log, 'w+')
import PyPDF2
pdf_reader = PyPDF2.PdfFileReader(pdf_file)
max_page = pdf_reader.numPages
for i in range(1, 3):
    pdf_writer = PyPDF2.PdfFileWriter()
    page = pdf_reader.getPage(i)
    page_string = page.extractText()
    idStart = int(page_string.find('Number'))+8
    idNumber = page_string[idStart: idStart+4]
    newPdf = idNumber+'.pdf'
    pdf_writer.addPage(page)
    pdf_idNumber = open(newPdf, 'wb')
    pdf_writer.write(pdf_idNumber)
    pdf_idNumber.close()
pdf_file.close()'''

yag = yagmail.SMTP("cambburnside@gmail.com", "Openfire1")
body = "For those who have lost the Cambridge log-in details, please see attached."
students = openpyxl.load_workbook('Camb11.xlsx')
max_value = students['Sheet1'].max_row

for row in range(2, max_value+1):
    school_ID_cell = 'B' + str(row)
    school_ID = students['Sheet1'][school_ID_cell].value
    address = str(school_ID)+'@burnside.school.nz'
    cambridge_ID_cell = 'C' + str(row)
    cambridge_ID = students['Sheet1'][cambridge_ID_cell].value
    filename = str(cambridge_ID)+'.pdf'
    print(school_ID, school_ID_cell, address, filename)
    yag.send(
        to=address,
        subject="Cambridge Results: Your log-in and password",
        contents=body,
        attachments=filename,
    )

