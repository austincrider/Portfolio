# OVERVIEW

* Queries the Cisco Unity server AXL API to download all user and phone extensions.
* The exported data is used to add phone extension data into Active Directory and Office365.

# PROGRAM PROCESS

1. A scheduled Windows task runs the batch file (getPhoneExtensions.bat)
2. The batch file queries the Cisco Unity API using an XML request sent to the AXL interface via curl
3. The results of the request are written to an XML file (destFile.xml)
3. The batch file starts a Python program to convert the XML file into a CSV file
4. The final CSV file is output C:\Automation\Output\phoneExtensions.csv
5. A separate ETL program can be used to processes the file and write to Active Directory attributes

# FILES

1. getPhoneExtensions.bat = the parent script
2. convertXmlToCsv.py = the XML to CSV conversion script 
3. destFile.xml = the output file from the API request
4. requests.xml = XML file for initiating the API request. Specifies what Unity data is needed.

# AUTHOR

**Austin Crider**
**2023-01-01**