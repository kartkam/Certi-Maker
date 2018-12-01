import sys
import os


#initializing important variables...
templateFileName = sys.argv[1]
excelFileName = sys.argv[2]
folderToSaveIn = sys.argv[3]


#create folder to store certificates
def createFolderForCertificates():
	if not os.path.exists(folderToSaveIn):
		os.makedirs(folderToSaveIn)


#for converting doc to pdf
import comtypes.client

def convertToPdf(fileName):
	
	print("Started creating "+fileName+".pdf...")
	
	wdFormatPDF = 17
	
	in_file = os.path.abspath(fileName+".docx")
	out_file = os.path.abspath(fileName+".pdf")

	word = comtypes.client.CreateObject("Word.Application")
	doc = word.Documents.Open(in_file)
	doc.SaveAs(out_file, FileFormat=wdFormatPDF)
	doc.Close()
	word.Quit()
	
	print(fileName+".pdf created")
	
	del doc,word

	
#remove the temporary created word file
def removeTemporaryWordFile(fileName):
	os.remove(fileName+".docx")


#for creating certificates from template
from mailmerge import MailMerge
from datetime import date
import pandas as pd

def createCertificatesFromTemplate():

	participantsData = pd.read_excel(excelFileName,sheet=0)
	participantsCount = len(participantsData.index)


	for index in range(0,participantsCount):
		
		document = MailMerge(templateFileName)
		
		participantName = participantsData.iloc[index]["Name"]
		participantDate = participantsData.iloc[index]["Completed on"]
		
		document.merge(fullname=participantName,
			date="{:%d-%b-%Y}".format(participantDate)
		)

		generalFilePathOfCertificate = folderToSaveIn+"/"+participantName
		
		document.write(generalFilePathOfCertificate+".docx")
		
		convertToPdf(generalFilePathOfCertificate)
		
		del document
		
		removeTemporaryWordFile(generalFilePathOfCertificate)


#start of code
createFolderForCertificates()

createCertificatesFromTemplate()

