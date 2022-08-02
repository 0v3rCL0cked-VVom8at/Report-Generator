# Report-Generator
Project converting a csv to excel, filters data, then generates a word doc
This program auto generates word file reports used for Patch Remediation in Deltra systems. It takes as input specific .csv files pulled from report manager
it will then convert the csv into an excel spreadsheet and then filter the data.
Once filtered and selected it will populate the data into a word table and a word document will then be generated with a predertermined format. 

This project served to enhance my python knowledge as well as alleviate the extended process previously used to generate these reports. 

Python libraries used in this project:
-Tkinter (for gui and user input)
-openpyxl (excel-python interaction)
-docx (word-python interaction)
-pandas (csv reader data frame generator)
