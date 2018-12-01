# Certi-Maker

Generate certificates from MS word template and excel data.

Thanks to http://pbpython.com/python-word-template.html, which has been a great help in developing this code.

-> Open start.bat in text editor mode (notepad or other), make necessary changes to first three lines and save it.

Currently the code supports only 2 fields fullname and date. The python code will extract each row of excel data and create a certificate in the mentioned folder path.

-> To use your own template: 

I. Before executing the code:
 
1. Create a certificate template in MS Word.
2. Navigate to Insert -> Quick parts -> Field...
3. Add a MergeField, add name of field, click OK.
4. Repeat the steps to add other fields.

II. In participants.xlsx, add/customize the necessary columns for the fields created.

III. In certiMaker.py, at line 63, change the document.merge method to include the added/customized new fields.

-> Double-click start.bat, a cmd screen appears, message will be displayed as each certificate gets created.