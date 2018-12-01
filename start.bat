set templateFileName=<file name of certificate template with extension>
set excelFileName=<excel file where data is present>
set folderToSaveIn=<name of folder to save the certificates>

python certiMaker.py %templateFileName% %excelFileName% %folderToSaveIn%

@pause