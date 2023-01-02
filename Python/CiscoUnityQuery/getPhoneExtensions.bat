@echo off

curl -k -u USERNAMEHERE:PASSWORDHERE -H "Content-type:text/xml;" -H 'SOAPAction: "CUCM:DB ver=14.0 executeSQLQuery"' -d @requests.xml https://10.1.10.10:8443/axl/ > destFile.xml

py convertXmlToCsv.py