import pandas as pd
from bs4 import BeautifulSoup

# Before first run, install these modules:
# python3 -m pip install beautifulsoup4
# python3 -m pip install pandas
# python3 -m pip install lxml

##### INITIAL SETUP
# Set max columns and display width for printing data frames to screen
pd.set_option("display.max_columns", 2)
pd.set_option("display.width", 1000)

##### READ XML FILE
xml_file = open('destFile.xml', 'r')
contents = xml_file.read()

# Create BS object
soup = BeautifulSoup(contents, 'html.parser')

# Extract data and assign it to lists
email = soup.find_all("mailid")
number = soup.find_all("telephonenumber")

##### LOOP THROUGH DATA
# Create list for appending data
phone_info = []

# Uses the length of one of the lists (email) to loop through the pages of Unity
for i in range(0, len(email)):
    rows = [email[i].get_text(),
            number[i].get_text()]
    phone_info.append(rows)

# Create a dataframe with Pandas and print
df = pd.DataFrame(phone_info, columns=['Email', 'Phone'])
print(df)

##### OUTPUT TO FILE
directory = df.to_csv("c:\\automation\\output\\phones\\phoneExtensions.csv", index=False)