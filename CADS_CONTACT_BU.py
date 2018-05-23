import pandas as pd
import numpy as np
import xlsxwriter


dd_source = pd.ExcelFile(r"C:\Users\eakarsu\Documents\Python\CADS\BU\Source Files\DD_CADS_Full.xls")
dd_sheet_1 = pd.read_excel(dd_source, 'Employee')
dd_sheet_2 = pd.read_excel(dd_source, 'Employee 2')

dd_employee = pd.concat([dd_sheet_1,dd_sheet_2]);
dd_employee = dd_employee[dd_employee['Organization'].str[:2] != 'TG']
dd_employee = dd_employee[dd_employee['Series/Grade'] != 'Other']
dd_employee = dd_employee[dd_employee['Series/Grade'] != 'Other (Intern)']

dd_employee['Empl ID'].fillna(value="Not Available", inplace=True)
dd_employee = dd_employee[dd_employee['Empl ID'] != 'Not Available']

dd_employee.rename(inplace=True, columns={"Empl ID": "LocalUserID"})

ldap_bu = dd_employee.copy()

''' CONTACT_BU_2 '''

ldap_bu2 = ldap_bu.copy()
ldap_bu2 = ldap_bu2[(ldap_bu2['Organization'].str[4:18] == '00000000000000')]

ldap_bu2['OrganizationName'] = ""
ldap_bu2['OrganizationName'] = np.where(ldap_bu2['Level 2'] != '' ,ldap_bu2['Level 2']+ '<br>'+ '<br>'+ '<br>'+ '<br>'+ '<br>'+ '<br>' , ldap_bu2['OrganizationName'] )
ldap_bu2['OrganizationName'].fillna(value="Not Available", inplace=True)
ldap_bu2 = ldap_bu2[ldap_bu2['OrganizationName'] != 'Not Available']

ldap_bu2.rename(inplace=True, columns={"Organization": "OrgCodeLevel2"})

#ldap_bu2['OrgCodeLevel2'] =  pd.to_numeric(ldap_bu2['OrgCodeLevel2'])

bu2_final = ldap_bu2[['LocalUserID','OrgCodeLevel2','OrganizationName']]

'''Save new Table and Export '''
ldap_bu_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\CONTACT_BU\Target Files\Contact_BU_2" + '.xlsx', engine='xlsxwriter')
bu2_final.to_excel(ldap_bu_writer, 'LDAP_2', index=False)
ldap_bu_writer.save()


''' CONTACT_BU_3 '''


ldap_bu3 = ldap_bu.copy()

ldap_bu3 = ldap_bu3[((ldap_bu3['Organization'].str[6:18] == '000000000000') & (ldap_bu3['Organization'].str[4:6] > '00')) | ((ldap_bu3['Organization'].str[4:18] == '00000000000001') & (ldap_bu3['Organization'].str[2:4] > '00')) ]

#| ((ldap_bu3['Organization'].str[4:18] == '00000000000001') & (ldap_bu3['Organization'].str[3:5] != '00'))]  
#ldap_bu3 = ldap_bu3[((ldap_bu3['Organization'].str[4:18] == '00000000000001')& (ldap_bu3['Organization'].str[2:4] > '00'))]

#ldap_bu3 = pd.concat([ldap_bu3_1,ldap_bu3_2]);


# this is necessary for allowing certain fields that have an executive level 8 but not current level name
ldap_bu3['Level 8'].fillna(value="", inplace=True)
ldap_bu3['Level 3'].fillna(value="", inplace=True)


ldap_bu3['OrganizationName'] = ""
ldap_bu3['OrganizationName'] = ldap_bu3.apply(lambda x: ldap_bu3['Level 2'] + '<br>' + ldap_bu3['Level 3'] + '<br>'+ '<br>'+ '<br>'+ '<br>'+ '<br>' + ldap_bu3['Level 8'])

ldap_bu3['OrganizationName'].fillna(value="Not Available", inplace=True)
#ldap_bu3 = ldap_bu3[(ldap_bu3['OrganizationName'] != 'Not Available') & (ldap_bu3['OrganizationName'] != "")]

ldap_bu3.rename(inplace=True, columns={"Organization": "OrgCodeLevel3"})
#ldap_bu3['OrgCodeLevel3'] =  pd.to_numeric(ldap_bu3['OrgCodeLevel3'])

bu3_final = ldap_bu3[['LocalUserID','OrgCodeLevel3','OrganizationName']]

'''Save new Table and Export '''
ldap_bu_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\CONTACT_BU\Target Files\Contact_BU_3" + '.xlsx', engine='xlsxwriter')
bu3_final.to_excel(ldap_bu_writer, 'LDAP_3', index=False)
ldap_bu_writer.save()


''' CONTACT 4 '''

ldap_bu4 = ldap_bu.copy()

ldap_bu4 = ldap_bu4[((ldap_bu4['Organization'].str[10:18] == '00000000') & (ldap_bu4['Organization'].str[6:10] > '0000')) | ((ldap_bu4['Organization'].str[6:18] == '000000000001') & (ldap_bu4['Organization'].str[4:6] > '00'))]  


ldap_bu4['Level 8'].fillna(value="", inplace=True)
ldap_bu4['Level 4'].fillna(value="", inplace=True)

ldap_bu4['OrganizationName'] = ""
ldap_bu4['OrganizationName'] = ldap_bu4.apply(lambda x: ldap_bu4['Level 2'] + '<br>' + ldap_bu4['Level 3'] + '<br>'+ ldap_bu4['Level 4'] + '<br>'+ '<br>'+ '<br>'+ '<br>' + ldap_bu4['Level 8'])

ldap_bu4['OrganizationName'].fillna(value="Not Available", inplace=True)
ldap_bu4 = ldap_bu4[(ldap_bu4['OrganizationName'] != 'Not Available') & (ldap_bu4['OrganizationName'] != "")]

ldap_bu4.rename(inplace=True, columns={"Organization": "OrgCodeLevel4"})

ldap_bu4['OrgCodeLevel4'] =  pd.to_numeric(ldap_bu4['OrgCodeLevel4'])

bu4_final = ldap_bu4[['LocalUserID','OrgCodeLevel4','OrganizationName']]

'''Save new Table and Export '''
ldap_bu_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\CONTACT_BU\Target Files\Contact_BU_4" + '.xlsx', engine='xlsxwriter')
bu4_final.to_excel(ldap_bu_writer, 'LDAP_4', index=False)
ldap_bu_writer.save()

''' CONTACT 5 '''
ldap_bu5 = ldap_bu.copy()

ldap_bu5 = ldap_bu5[((ldap_bu5['Organization'].str[12:18] == '000000') & (ldap_bu5['Organization'].str[10:12] > '00'))
 | ((ldap_bu5['Organization'].str[10:18] == '00000001') & (ldap_bu5['Organization'].str[6:10] > '0000'))]  

ldap_bu5['Level 8'].fillna(value="", inplace=True)
ldap_bu5['Level 5'].fillna(value="", inplace=True)

ldap_bu5['OrganizationName'] = ""
ldap_bu5['OrganizationName'] = ldap_bu5.apply(lambda x: ldap_bu5['Level 2'] + '<br>' + ldap_bu5['Level 3'] + '<br>'+ ldap_bu5['Level 4'] + '<br>'+ ldap_bu5['Level 5'] + '<br>'+ '<br>'+ '<br>' + ldap_bu5['Level 8'])

ldap_bu5['OrganizationName'].fillna(value="Not Available", inplace=True)
ldap_bu5 = ldap_bu5[(ldap_bu5['OrganizationName'] != 'Not Available') & (ldap_bu5['OrganizationName'] != "")]

ldap_bu5.rename(inplace=True, columns={"Organization": "OrgCodeLevel5"})
ldap_bu5['OrgCodeLevel5'] =  pd.to_numeric(ldap_bu5['OrgCodeLevel5'])

bu5_final = ldap_bu5[['LocalUserID','OrgCodeLevel5','OrganizationName']]

'''Save new Table and Export '''
ldap_bu_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\CONTACT_BU\Target Files\Contact_BU_5" + '.xlsx', engine='xlsxwriter')
bu5_final.to_excel(ldap_bu_writer, 'LDAP_5', index=False)
ldap_bu_writer.save()

''' CONTACT 6 '''
ldap_bu6 = ldap_bu.copy()

ldap_bu6 = ldap_bu6[((ldap_bu6['Organization'].str[14:18] == '0000') & (ldap_bu6['Organization'].str[12:14] > '00'))
 | ((ldap_bu6['Organization'].str[12:18] == '000001') & (ldap_bu6['Organization'].str[10:12] > '00'))]  

ldap_bu6['Level 8'].fillna(value="", inplace=True)
ldap_bu6['Level 6'].fillna(value="", inplace=True)

ldap_bu6['OrganizationName'] = ""
ldap_bu6['OrganizationName'] = ldap_bu6.apply(lambda x: ldap_bu6['Level 2'] + '<br>' + ldap_bu6['Level 3'] + '<br>'+ ldap_bu6['Level 4'] + '<br>'+ ldap_bu6['Level 5'] + '<br>'+ ldap_bu6['Level 6']+ '<br>'+ '<br>' + ldap_bu6['Level 8'])

ldap_bu6['OrganizationName'].fillna(value="Not Available", inplace=True)
ldap_bu6 = ldap_bu6[(ldap_bu6['OrganizationName'] != 'Not Available') & (ldap_bu6['OrganizationName'] != "")]

ldap_bu6.rename(inplace=True, columns={"Organization": "OrgCodeLevel6"})
ldap_bu6['OrgCodeLevel6'] =  pd.to_numeric(ldap_bu6['OrgCodeLevel6'])

bu6_final = ldap_bu6[['LocalUserID','OrgCodeLevel6','OrganizationName']]

'''Save new Table and Export '''
ldap_bu_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\CONTACT_BU\Target Files\Contact_BU_6" + '.xlsx', engine='xlsxwriter')
bu6_final.to_excel(ldap_bu_writer, 'LDAP_6', index=False)
ldap_bu_writer.save()

''' CONTACT 7 '''
ldap_bu7 = ldap_bu.copy()

ldap_bu7 = ldap_bu7[(ldap_bu7['Organization'].str[16:18] == '00') & (ldap_bu7['Organization'].str[14:16] > '00')]  

ldap_bu7['Level 8'].fillna(value="", inplace=True)
ldap_bu7['Level 7'].fillna(value="", inplace=True)

ldap_bu7['OrganizationName'] = ""
ldap_bu7['OrganizationName'] = ldap_bu7.apply(lambda x: ldap_bu7['Level 2'] + '<br>' + ldap_bu7['Level 3'] + '<br>'+ ldap_bu7['Level 4'] + '<br>'+ ldap_bu7['Level 5'] + '<br>'+ ldap_bu7['Level 6']+ '<br>'+ ldap_bu7['Level 7']+ '<br>' + ldap_bu7['Level 8'])

ldap_bu7['OrganizationName'].fillna(value="Not Available", inplace=True)
ldap_bu7 = ldap_bu7[(ldap_bu7['OrganizationName'] != 'Not Available') & (ldap_bu7['OrganizationName'] != "")]

ldap_bu7.rename(inplace=True, columns={"Organization": "OrgCodeLevel7"})
ldap_bu7['OrgCodeLevel7'] =  pd.to_numeric(ldap_bu7['OrgCodeLevel7'])

bu7_final = ldap_bu7[['LocalUserID','OrgCodeLevel7','OrganizationName']]

'''Save new Table and Export '''
ldap_bu_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\CONTACT_BU\Target Files\Contact_BU_7" + '.xlsx', engine='xlsxwriter')
bu7_final.to_excel(ldap_bu_writer, 'LDAP_7', index=False)
ldap_bu_writer.save()

''' CONTACT 8 '''
ldap_bu8 = ldap_bu.copy()

ldap_bu8 = ldap_bu8[(ldap_bu8['Organization'].str[14:16] > '00') & (ldap_bu8['Organization'].str[15:17] > '00')]  

ldap_bu8['OrganizationName'] = ""
ldap_bu8['OrganizationName'] = ldap_bu8.apply(lambda x: ldap_bu8['Level 2'] + '<br>' + ldap_bu8['Level 3'] + '<br>'+ ldap_bu8['Level 4'] + '<br>'+ ldap_bu8['Level 5'] + '<br>'+ ldap_bu8['Level 6']+ '<br>'+ ldap_bu8['Level 7']+ '<br>' + ldap_bu8['Level 8'])

ldap_bu8['OrganizationName'].fillna(value="Not Available", inplace=True)
ldap_bu8 = ldap_bu8[(ldap_bu8['OrganizationName'] != 'Not Available') & (ldap_bu8['OrganizationName'] != "")]

ldap_bu8.rename(inplace=True, columns={"Organization": "OrgCodeLevel8"})
ldap_bu8['OrgCodeLevel8'] =  pd.to_numeric(ldap_bu8['OrgCodeLevel8'])

bu8_final = ldap_bu8[['LocalUserID','OrgCodeLevel8','OrganizationName']]

'''Save new Table and Export '''
ldap_bu_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\CONTACT_BU\Target Files\Contact_BU_8" + '.xlsx', engine='xlsxwriter')
bu8_final.to_excel(ldap_bu_writer, 'LDAP_8', index=False)
ldap_bu_writer.save()


all_contact_bu_levels = pd.concat([bu2_final, bu3_final, bu4_final, bu5_final, bu6_final, bu7_final, bu8_final]);

all_contacts = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\CONTACT_BU\Target Files\ALL_CONTACT_BU" + '.xlsx', engine='xlsxwriter')
all_contact_bu_levels.to_excel(all_contacts, 'LDAP_8', index=False)
all_contacts.save()