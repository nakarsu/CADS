
import pandas as pd
import numpy as np
import xlsxwriter
from datetime import date



people_file = pd.read_excel(r"C:\Users\eakarsu\Documents\Python\CADS\PEOPLE_MASTER\Source Files\People_new.xlsx")

dd_source = pd.ExcelFile(r"C:\Users\eakarsu\Documents\Python\CADS\PEOPLE_MASTER\Source Files\DD_CADS_Full.xls")
dd_sheet_1 = pd.read_excel(dd_source, 'Employee')
dd_sheet_2 = pd.read_excel(dd_source, 'Employee 2')
dd_employee = pd.concat([dd_sheet_1,dd_sheet_2])

dd_employee = dd_employee[dd_employee['Organization'].str[:2] != 'TG']
dd_employee = dd_employee[dd_employee['Series/Grade'] != 'Other']
dd_employee = dd_employee[dd_employee['Series/Grade'] != 'Other (Intern)']
dd_employee['Empl ID'].fillna(value="Not Available", inplace=True)
dd_employee = dd_employee[dd_employee['Empl ID'] != 'Not Available']


del people_file['dn']
#del people_file['organizationUnitDN']
#del people_file['DepartmentNumber']
people_file = people_file[(people_file['businessCategory'] == 'EMPLOYEE') | (people_file['businessCategory'] == 'CONTRACTOR')]

#Format UI based on the first two digits (ds,ci,prod)
people_file['uid'] = np.where(people_file['uid'].str[:2] == "ds", people_file['uid'].str[:3]+ "\\" + people_file['uid'].str[3:11], people_file['uid'])
people_file['uid'] = np.where(people_file['uid'].str[:2] == "ci", people_file['uid'].str[:3]+ "\\" + people_file['uid'].str[3:11], people_file['uid'])
people_file['uid'] = np.where(people_file['uid'].str[:2] == "prod", people_file['uid'].str[:5]+ "\\" + people_file['uid'].str[5:13], people_file['uid'])


#filter n/a from each column
people_file['initials'].replace('N/A', "", inplace=True)
people_file['uid'].replace('N/A', "", inplace=True)
people_file['payPlan'].replace('N/A', "", inplace=True)
people_file['payGrade'].replace('N/A', "", inplace=True)
people_file['userClass'].replace('N/A', "", inplace=True)
people_file['mail'].replace('N/A', "", inplace=True)
people_file['street'].replace('N/A', "", inplace=True)
people_file['title'].replace('N/A', "", inplace=True)
people_file['l'].replace('N/A', "", inplace=True)
people_file['st'].replace('N/A', "", inplace=True)
people_file['postalCode'].replace('N/A', "", inplace=True)
people_file['postalCode'].replace(0, "", inplace=True)
people_file['mobile'].replace('N/A', "", inplace=True)
people_file['ou'].replace('N/A', "", inplace=True)
people_file['houseIdentifier'].replace('N/A', "", inplace=True)

#Add Organization Name column
people_file['OrganizationName'] = ""

#Merge Discovery Data with People_file to extrraact the neceasarry data for newly created columns above.
#Curently going to merge with two keys Uniqueidentifier and Local User ID
'''
people_merged = pd.merge(people_file,
dd_employee[['SEID','Empl ID', 'M SEID' ,'Manager',"Manager's Phone","Manager's Mobile","Manager's E-Mail",'Organization','Level 2','Level 3','Level 4','Level 5','Level 6','Level 7','Level 8']],
left_on='ou', right_on='SEID', how="left")
'''

people_merge_1 = pd.merge(people_file,
dd_employee[['SEID','Empl ID']],
left_on='ou', right_on='SEID', how="left")
'''
people_merge_1['localUserId'].fillna(value="Not Available", inplace=True)
people_merge_1 = people_merge_1[people_merge_1['localUserId'] != 'Not Available']
'''
people_merge_1.rename(inplace=True, columns={"SEID": "SEIDx"})

people_merged = pd.merge(people_merge_1,
dd_employee[['SEID', 'Manager',"Manager's Phone","Manager's Mobile","Manager's E-Mail",'Organization','Level 2','Level 3','Level 4','Level 5','Level 6','Level 7','Level 8']],
left_on='uniqueIdentifier', right_on='SEID', how="left")

# Only keep the rows that contain the localUserID in discovery data delete the ones that are not found in dd
#people_merged = people_merged[people_merged['localUserId'].isin(dd_employee['Empl ID'])]


people_merged.rename(inplace=True, columns={"Empl ID": "ManagerEMPLID"})
people_merged.rename(inplace=True, columns={"Manager": "Mgr_name"})
people_merged.rename(inplace=True, columns={"Manager's Phone": "Mgr_Phone"})
people_merged.rename(inplace=True, columns={"Manager's Mobile": "Mgr_Mobile"})
people_merged.rename(inplace=True, columns={"Manager's E-Mail": "Mgr_EMail"})
people_merged.rename(inplace=True, columns={"Organization": "OrganizationCode"})


people_merged['Mgr_Phone'].replace('N/A', "", inplace=True)
people_merged['Mgr_Mobile'].replace('N/A', "", inplace=True)
people_merged['Mgr_EMail'].replace('N/A', "", inplace=True)
people_merged['ManagerEMPLID'].replace('N/A', "", inplace=True)
people_merged['Mgr_name'].replace('N/A', "", inplace=True)
people_merged['Mgr_Phone'].fillna(value="", inplace=True)
people_merged['Mgr_Mobile'].fillna(value="", inplace=True)

people_merged['telephone'].fillna(value="", inplace=True)
people_merged['mobile'].fillna(value="", inplace=True)

#Convert Phone and Mobile Numbers to string first
people_merged['telephone'] = people_merged['telephone'].astype(str)
people_merged['mobile'] = people_merged['mobile'].astype(str)
people_merged['Mgr_Phone'] = people_merged['Mgr_Phone'].astype(str)
people_merged['Mgr_Mobile'] = people_merged['Mgr_Mobile'].astype(str)

#Format the Phone numbers
people_merged['telephone'] = np.where(people_merged['telephone'] != "", people_merged['telephone'].str[:3]+people_merged['telephone'].str[4:7]+people_merged['telephone'].str[8:12], people_merged['telephone'])
people_merged['mobile'] = np.where(people_merged['mobile'].str[:2] == "+1", people_merged['mobile'].str[3:6]+people_merged['mobile'].str[7:10]+people_merged['mobile'].str[11:15], people_merged['mobile'])
people_merged['Mgr_Phone'] = np.where(people_merged['Mgr_Phone'] != "", people_merged['Mgr_Phone'].str[:3]+people_merged['Mgr_Phone'].str[4:7]+people_merged['Mgr_Phone'].str[8:12], people_merged['Mgr_Phone'])
people_merged['Mgr_Mobile'] = np.where(people_merged['Mgr_Mobile'] != "", people_merged['Mgr_Mobile'].str[:3]+people_merged['Mgr_Mobile'].str[4:7]+people_merged['Mgr_Mobile'].str[8:12], people_merged['Mgr_Mobile'])

#Convert Phone numbers from String to Number 
people_merged[['telephone']] = people_merged[['telephone']].apply(pd.to_numeric)
people_merged[['mobile']] = people_merged[['mobile']].apply(pd.to_numeric)
people_merged[['Mgr_Phone']] = people_merged[['Mgr_Phone']].apply(pd.to_numeric)
people_merged[['Mgr_Mobile']] = people_merged[['Mgr_Mobile']].apply(pd.to_numeric)

#Fill #N/A for each null value of ManagerEMPLID
people_merged['ManagerEMPLID'].fillna(value="#N/A", inplace=True)

#Convert Date Time 
people_merged['createTimeStamp'] = people_merged.apply(lambda x: people_merged['createTimeStamp'].str[:12])
people_merged['createTimeStamp'] = people_merged['createTimeStamp'].apply(lambda x: pd.to_datetime(str(x), format='%Y%m%d%H%M'))

people_merged['modifyTimeStamp'] = people_merged.apply(lambda x: people_merged['modifyTimeStamp'].str[:12])
people_merged['modifyTimeStamp'] = people_merged['modifyTimeStamp'].apply(lambda x: pd.to_datetime(str(x), format='%Y%m%d%H%M'))

#Necessary for concatenation of null values as blank spaces
people_merged['Level 2'].fillna(value="", inplace=True)
people_merged['Level 3'].fillna(value="", inplace=True)
people_merged['Level 4'].fillna(value="", inplace=True)
people_merged['Level 5'].fillna(value="", inplace=True)
people_merged['Level 6'].fillna(value="", inplace=True)
people_merged['Level 7'].fillna(value="", inplace=True)
people_merged['Level 8'].fillna(value="", inplace=True)

people_merged['OrganizationName'] = people_merged.apply(lambda x: people_merged['Level 2'] + '<br>' + people_merged['Level 3'] + '<br>'+ people_merged['Level 4'] + '<br>'+ people_merged['Level 5'] + '<br>' + people_merged['Level 6']+ '<br>'+ people_merged['Level 7']+ '<br>' +  people_merged['Level 8'])

#people_merged = people_merged[['localUserId','uniqueIdentifier','ManagerEMPLID','Mgr_name']]
people_merged = people_merged[['localUserId','uniqueIdentifier','cn','sn','givenName','initials','uid','businessCategory','payPlan','payGrade','mail','title','employeeType','street','l','st','postalCode','telephone','mobile','ManagerEMPLID','Mgr_name','Mgr_Phone','Mgr_Mobile','Mgr_EMail','houseIdentifier','userClass','createTimeStamp','modifyTimeStamp','OrganizationCode','OrganizationName']]


'''Save new Table and Export '''
people_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\PEOPLE_MASTER\Target Files\CADS_People_Master" + '.xlsx', engine='xlsxwriter')
people_merged.to_excel(people_writer, 'LDAP_2', index=False)
people_writer.save()