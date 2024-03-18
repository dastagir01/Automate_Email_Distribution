import smtplib
import pandas
import openpyxl

my_email = <INPUT YOUR EMAIL>
password = <INPUT YOUR EMAIL PASSWORD>

# Obtain Subcontractor Database
data = pandas.read_excel('subcontractor_db.xlsx', sheet_name='Subcontractors')
#convert data to dictionary; orient="records" places it in a row
data_dict = data.to_dict(orient="records")

# Initialize Project Data
project_data = pandas.read_excel('subcontractor_db.xlsx', sheet_name='Project_Info')

COMPANY_NAME = project_data.iloc[0, 1]
PROJECT_NAME = project_data.iloc[1, 1]
PROJECT_CITY = project_data.iloc[2, 1]
PROJECT_STATE = project_data.iloc[3, 1]
NUM_UNITS = project_data.iloc[4, 1]
NUM_STORIES = project_data.iloc[5, 1]
PHASE = project_data.iloc[6, 1]
SENDER_NAME = project_data.iloc[8, 1]
SENDER_TILE = project_data.iloc[9, 1]
NUM_UNITS = str(NUM_UNITS)
NUM_STORIES = str(NUM_STORIES)

# Send mass email to subcontractors
for sub in data_dict:
    Followup = sub['Send Email']
    Trade = sub['Trades']
    Company = sub['Company']
    First_Name = sub['First Name']
    Email = sub['Email']

    #get letter file
    file_path = f"letter_templates/letter_1.txt"
    # below print statement to confirm code is working properly
    print(f"Current Trade is:{Trade} & Subcontractor is: {Company}")
    with open(file_path) as letter_file:
        #read letter file
        content = letter_file.read()
        #Replace letter content
        content = content.replace("[NAME]", First_Name)
        content = content.replace("[COMPANY]", Company)
        content = content.replace("[SENDER_NAME]", SENDER_NAME)
        content = content.replace("[COMPANY_NAME]", COMPANY_NAME)
        content = content.replace("[PROJECT_NAME]", PROJECT_NAME)
        content = content.replace("[PROJECT_CITY]", PROJECT_CITY)
        content = content.replace("[PROJECT_STATE]", PROJECT_STATE)
        content = content.replace("[PHASE]", PHASE)
        content = content.replace("[NUM_UNITS]", NUM_UNITS)
        content = content.replace("[NUM_STORIES]", NUM_STORIES)
        content = content.replace("[Trade]", Trade)
        content = content.replace("[SENDER_TITLE]", SENDER_TILE)

        with smtplib.SMTP("smtp.gmail.com") as connection:
            #transport layer security(TLS) (creating secure/encrypted connection)
            connection.starttls()
            #loggin in with email & password
            connection.login(user=my_email, password=password)
            connection.sendmail(from_addr=my_email,to_addrs= Email, msg=f"Subject:{PROJECT_NAME}- {Trade} - {Company}\n\n {content}")
