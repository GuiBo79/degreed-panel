import pandas as pd
from datetime import date
import os
import subprocess

print("    __ DEGREED __")
print("   /.-'       `-.\ ")
print("  j/             \j ")
print(" j/_______________\j ")
print("/o.-==-. .-. .-==-.o\ ")
print("||      )) ((      || ")
print(" \\____//    j\____// ")
print("  `-==-'     `-==-' ")

print("  _____       _ _ _                                ____             _        _                 ")
print(" / ____|     (_) | |                              |  _ \           | |      | |                ")
print("| |  __ _   _ _| | |__   ___ _ __ _ __ ___   ___  | |_) | ___  _ __| |_ ___ | | __ _ ___  ___  ")
print("| | |_ | | | | | | '_ \ / _ \ '__| '_ ` _ \ / _ \ |  _ < / _ \| '__| __/ _ \| |/ _` / __|/ _ \ ")
print("| |__| | |_| | | | | | |  __/ |  | | | | | |  __/ | |_) | (_) | |  | || (_) | | (_| \__ \ (_) |")
print(" \_____|\__,_|_|_|_| |_|\___|_|  |_| |_| |_|\___| |____/ \___/|_|   \__\___/|_|\__,_|___/\___/ ")
                                                                                                
                                                                                                
print("Guilherme Bortolaso 2020")


data = pd.read_excel (r'C:\Users\guibo\OneDrive\Área de Trabalho\Degreed\User_file template_TIC.xlsx', sheet_name='User List')
rules = pd.read_excel (r'C:\Users\guibo\OneDrive\Área de Trabalho\Degreed\BusinessRules Template_TIC.xlsx', sheet_name='Business Rules')

for i in range(len(data)): ## Updating Users
    if data["Delete"][i] == "Y":
        data = data.drop([i])
    
for i in range(len(rules)): ## Updating Rules
    if rules["Delete"][i] == "Y":
        rules = rules.drop([i])

bu_unique = data["Business Unit"].unique() ## Getting Unique Business Units

locations_unique = data["Office Location"].unique() ## Getting Unique Locations

locations_unique = list(locations_unique)

regions = {'Sao Paulo, Brazil': 'LATAM',           ## Hard Coded Regions Tables
           'Mexico City, Mexico': 'NAMER',
           'New York, NY': 'NAMER',
           'Toronto, Canada': 'NAMER',
           'Bangalore, India': 'APAC',
           'Los Angeles, CA': 'NAMER',
           'Chicago, IL': 'NAMER',
           'Buenos Aires, Argentina': 'LATAM',
           'Beijing, China': 'APAC',
           'Tokyo, Japan': 'APAC',
           'London, UK': 'EMEA'}

def get_region(location):     ## Return Regio by Location
    for key,value in regions.items():
        if key == location:
            return value

users_file = list(data.head()) ## Constructing the File Format
users_file.append("All Hands")
for bu in bu_unique:
    users_file.append(bu)
users_file.append("Leadership Dev")
users_file.append("APAC")
users_file.append("EMEA")
users_file.append("LATAM")
users_file.append("NAMER")

############# Rules parse function from Excel ######################

def assert_rules_1(data):
    for if_s, opt_1, equals_1 in zip(rules['if'],rules['op 1'],rules['equals 1']): ## Assertion coverting from Excel
        if opt_1 == "=":
            if data[if_s] == equals_1:
                assertion1.append(True)
            else:
                assertion1.append(False)
        elif opt_1 == ">":
            if data[if_s] > equals_1:
                assertion1.append(True)
            else:
                assertion1.append(False)
        elif opt_1 == "<":
            if data[if_s] < equals_1:
                assertion1.append(True)
            else:
                assertion1.append(False)
        elif opt_1 == "<=":
            if data[if_s] <= equals_1:
                assertion1.append(True)
            else:
                assertion1.append(False)
        elif opt_1 == ">=":
            if data[if_s] >= equals_1:
                assertion1.append(True)
            else:
                assertion1.append(False)
                
def assert_rules_2(data):
    for if_s, opt_2, equals_2 in zip(rules['and if 2'],rules['op 2'],rules['equals 2']): ## Assertion coverting from Excel
        if opt_2 == "=":
            if data[if_s] == equals_2:
                assertion2.append(True)
            else:
                assertion2.append(False)
        elif opt_2 == ">":
            if data[if_s] > equals_2:
                assertion2.append(True)
            else:
                assertion2.append(False)
        elif opt_2 == "<":
            if data[if_s] < equals_2:
                assertion2.append(True)
            else:
                assertion2.append(False)
        elif opt_2 == "<=":
            if data[if_s] <= equals_2:
                assertion2.append(True)
            else:
                assertion2.append(False)
        elif opt_2 == ">=":
            if data[if_s] >= equals_2:
                assertion2.append(True)
            else:
                assertion2.append(False)
        else:
            assertion2.append(True)

print("---> Populating CSV File")
############# Populating Dataframe with User Data
csv = pd.DataFrame(columns = users_file)

for i in range(len(data)):
    assertion1 = []
    assertion2 = []
    assert_rules_1(data.iloc[i])
    assert_rules_2(data.iloc[i])

    
    group_matrix = []
    for assert_1,assert_2 in zip(assertion1,assertion2):
        if assert_1 and assert_2:
            group_matrix.append("Y")
        else:
            group_matrix.append("N")
            
    temp_data=[]
    for temp in data.iloc[i]:
        temp_data.append(temp)
        
        
    for j,rule in zip(range(len(rules)),rules["Groups"]):
        if rule == "All Hands":
            temp_data.append(group_matrix[j])
    for j,rule in zip(range(len(rules)),rules["Groups"]):
        if rule == "Finance":
            temp_data.append(group_matrix[j])
    for j,rule in zip(range(len(rules)),rules["Groups"]):
        if rule == "Sales":
            temp_data.append(group_matrix[j])
    for j,rule in zip(range(len(rules)),rules["Groups"]):
        if rule == "Production":
            temp_data.append(group_matrix[j])
    for j,rule in zip(range(len(rules)),rules["Groups"]):
        if rule == "HR":
            temp_data.append(group_matrix[j])
    for j,rule in zip(range(len(rules)),rules["Groups"]):
        if rule == "Marketing":
            temp_data.append(group_matrix[j])
    for j,rule in zip(range(len(rules)),rules["Groups"]):
        if rule == "Leadership Dev":
            temp_data.append(group_matrix[j])
    
   
            
    if get_region(data["Office Location"][i]) == "LATAM":
        temp_data.extend(("N","N","Y","N"))
    elif get_region(data["Office Location"][i]) == "NAMER":
        temp_data.extend(("N","N","N","Y"))
    elif get_region(data["Office Location"][i]) == "APAC":
        temp_data.extend(("Y","N","N","N"))
    elif get_region(data["Office Location"][i]) == "EMEA":
        temp_data.extend(("N","Y","N","N"))
        
    csv.loc[len(csv)] = temp_data

#### Saving Updated Files and CSV Formatted 

print("---> Saving Files")

csv_file = "User_File_Acme_"+date.today().strftime('%Y%m%d')+'.csv'
csv.to_csv(csv_file,index=False) ##CSV file to Server
csv.to_excel("User_file template_TIC_update_"+date.today().strftime('%Y%m%d')+ '.xlsx', index = False) ##Updated User Excel
rules.to_excel("BusinessRules Template_TIC_update_"+date.today().strftime('%Y%m%d')+ '.xlsx', index = False) ##Updated Rules Excel

##### Sending CSV File User Data to AWS Linux Server

print("---> Connecting to SSH and Sending Files to AWS Degreed Server")

pem = r"c:\Users\guibo\OneDrive\Área de Trabalho\Degreed\LightsailDefaultKey-us-east-1.pem"
csv2server = 'c:\\Users\\guibo\\OneDrive\\Área de Trabalho\\Degreed\\'  + csv_file
host = 'ec2-user@3.235.6.48:~'
subprocess.run(["scp","-i", pem, csv2server, host])

print("---> Process Successfully Concluded")






