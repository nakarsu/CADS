import pandas as pd
import numpy as np
import xlsxwriter
from datetime import date

#converters={'uniqueIdentifier':str}
#ldap_cads_bu = pd.read_excel(r"C:\Users\eakarsu\Documents\Python\CADS\BU\Source Files\LDAP_CADS_BU.txt")

ldap_cads_bu = pd.read_csv(r"C:\Users\eakarsu\Documents\Python\CADS\BU\Source Files\LDAP_CADS_BU.txt",sep='\t')

dd_source = pd.ExcelFile(r"C:\Users\eakarsu\Documents\Python\CADS\BU\Source Files\DD_CADS_Full.xls")
dd_sheet_1 = pd.read_excel(dd_source, 'Employee')
dd_sheet_2 = pd.read_excel(dd_source, 'Employee 2')
dd_employee = pd.concat([dd_sheet_1,dd_sheet_2])


dd_employee = dd_employee[dd_employee['Organization'].str[:2] != 'TG']
dd_employee = dd_employee[dd_employee['Series/Grade'] != 'Other']
dd_employee = dd_employee[dd_employee['Series/Grade'] != 'Other (Intern)']


ldap_cads_bu['createTimeStamp'] = ldap_cads_bu.apply(lambda x: ldap_cads_bu['createTimeStamp'].str[:12])
ldap_cads_bu['createTimeStamp'] = ldap_cads_bu['createTimeStamp'].apply(lambda x: pd.to_datetime(str(x), format='%Y%m%d%H%M'))

ldap_cads_bu['modifyTimeStamp'] = ldap_cads_bu.apply(lambda x: ldap_cads_bu['modifyTimeStamp'].str[:12])
ldap_cads_bu['modifyTimeStamp'] = ldap_cads_bu['modifyTimeStamp'].apply(lambda x: pd.to_datetime(str(x), format='%Y%m%d%H%M'))

ldap_cads_bu['Mgr_SEID'] = ""
#ldap_cads_bu['Mgr_SEID'] = np.where(ldap_cads_bu['manager']!="", ldap_cads_bu['manager'].str[17:22], ldap_cads_bu['Mgr_SEID'])
ldap_cads_bu['Mgr_SEID'] = ldap_cads_bu.apply(lambda x: ldap_cads_bu['manager'].str[17:22])
ldap_cads_bu['Mgr_SEID'].fillna(value="N/A", inplace=True)

del ldap_cads_bu['manager']

ldap_cads_bu['A'], ldap_cads_bu['OrgName'], ldap_cads_bu['Parent'] , ldap_cads_bu['D']  = ldap_cads_bu['dn'].str.split('=',3).str

del ldap_cads_bu['A']
del ldap_cads_bu['dn']
del ldap_cads_bu['D']

ldap_cads_bu['OrgName'], ldap_cads_bu['e1'] = ldap_cads_bu['OrgName'].str.split(',',1).str
ldap_cads_bu['Parent'], ldap_cads_bu['e2'] = ldap_cads_bu['Parent'].str.split(',',1).str

del ldap_cads_bu['e1']
del ldap_cads_bu['e2']


ldap_cads_bu.rename(inplace=True, columns={"uniqueIdentifier": "Organization"})
ldap_cads_bu.rename(inplace=True, columns={"ou": "OrgNameOriginal"})


ldap_cads_bu['Organization'] = ldap_cads_bu['Organization'].astype(str)

ldap_cads_bu = pd.merge(ldap_cads_bu,
dd_employee[['SEID', 'M SEID','Empl ID']],
left_on='Mgr_SEID', right_on='SEID', how="left")

ldap_cads_bu.rename(inplace=True, columns={"Empl ID": "ManagerEMPLID"})
del ldap_cads_bu['SEID']
del ldap_cads_bu['M SEID']
del ldap_cads_bu['Mgr_SEID']
del ldap_cads_bu['OrgName']

bu_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\BU\Target Files\CADS_BU_Delimeted" + '.xlsx', engine='xlsxwriter')
ldap_cads_bu.to_excel(bu_writer, 'cads_bu_merged', index=False)
bu_writer.save()

#This is were I will map the organization codes to buisness unit levels

'''Org Code Level 2'''

ldap_org_2 = ldap_cads_bu.copy()

ldap_org_2 = ldap_org_2[(ldap_org_2['Organization'].str[4:18] == '00000000000000')]

ldap_org_2['OrgCodeLevel2'] = ""
ldap_org_2['MajorBusinessUnitLevel2'] = ""
ldap_org_2['Bureau'] = "IRS"
ldap_org_2['OrgCodeLevel2'] = ldap_org_2.apply(lambda x: ldap_org_2['Organization'])
ldap_org_2['MajorBusinessUnitLevel2'] = ldap_org_2.apply(lambda x: ldap_org_2['OrgNameOriginal'])

#ldap_org_2 = ldap_org_2.drop_duplicates('Mgr_SEID')
ldap_org_2 = ldap_org_2[~ldap_org_2['MajorBusinessUnitLevel2'].str.contains('Exec')]

ldap_org_2['OrgCodeLevel2'] =  pd.to_numeric(ldap_org_2['OrgCodeLevel2'])

ldap_org_2 = ldap_org_2[['OrgCodeLevel2','MajorBusinessUnitLevel2','Bureau','createTimeStamp','modifyTimeStamp','ManagerEMPLID']]


org_2_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\BU\Target Files\LDAP_BU_2" + '.xlsx', engine='xlsxwriter')
ldap_org_2.to_excel(org_2_writer, 'cads_bu_merged', index=False)
org_2_writer.save()

'''Org Code Level 3'''

ldap_org_3 = ldap_cads_bu.copy()

ldap_org_3 = ldap_org_3[((ldap_org_3['Organization'].str[6:18] == '000000000000') & (ldap_org_3['Organization'].str[4:6] > '00')) | ((ldap_org_3['Organization'].str[4:18] == '00000000000001') & (ldap_org_3['Organization'].str[2:4] > '00')) ]

#ldap_org_3 = ldap_org_3[(ldap_org_3['Organization'].str[4:18] == '00000000000001') & (ldap_org_3['Organization'].str[2:4] > '00')]  

#Using the Parent Organization name as MajorBusinessUnitReference
ldap_org_3['MajorBusinessUnitReference'] = ""
ldap_org_3['MajorBusinessUnitReference'] = ldap_org_3.apply(lambda x: ldap_org_3['Parent'])

ldap_org_3['ParentCodeLevel2'] = ''
ldap_org_3['ParentCodeLevel2'] = ldap_org_3.apply(lambda x: ldap_org_3['Organization'].str[:4] + '00000000000000')

ldap_org_3['OrgCodeLevel3'] = ""
ldap_org_3['OrgNameLevel3'] = ""
ldap_org_3['OrgCodeLevel3'] = ldap_org_3.apply(lambda x: ldap_org_3['Organization'])
ldap_org_3['OrgNameLevel3'] = ldap_org_3.apply(lambda x: ldap_org_3['OrgNameOriginal'])

ldap_org_3['OrgCodeLevel3'] =  pd.to_numeric(ldap_org_3['OrgCodeLevel3'])
ldap_org_3['ParentCodeLevel2'] =  pd.to_numeric(ldap_org_3['ParentCodeLevel2'])

ldap_org_3 = ldap_org_3[['OrgCodeLevel3','OrgNameLevel3','MajorBusinessUnitReference','ParentCodeLevel2','createTimeStamp','modifyTimeStamp','ManagerEMPLID']]


org_3_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\BU\Target Files\LDAP_BU_3" + '.xlsx', engine='xlsxwriter')
ldap_org_3.to_excel(org_3_writer, 'cads_bu_merged', index=False)
org_3_writer.save()

'''Org Code Level 4'''

ldap_org_4 = ldap_cads_bu.copy()

ldap_org_4 = ldap_org_4[((ldap_org_4['Organization'].str[10:18] == '00000000') & (ldap_org_4['Organization'].str[6:10] > '0000')) | ((ldap_org_4['Organization'].str[6:18] == '000000000001') & (ldap_org_4['Organization'].str[4:6] > '00'))]  

#Using the Parent Organization name as MajorBusinessUnitReference
ldap_org_4['Level3Reference'] = ""
ldap_org_4['Level3Reference'] = ldap_org_4.apply(lambda x: ldap_org_4['Parent'])

ldap_org_4['ParentCodeLevel3'] = ''
ldap_org_4['ParentCodeLevel3'] = ldap_org_4.apply(lambda x: ldap_org_4['Organization'].str[:6] + '000000000000')

ldap_org_4['OrgCodeLevel4'] = ""
ldap_org_4['OrgNameLevel4'] = ""
ldap_org_4['OrgCodeLevel4'] = ldap_org_4.apply(lambda x: ldap_org_4['Organization'])
ldap_org_4['OrgNameLevel4'] = ldap_org_4.apply(lambda x: ldap_org_4['OrgNameOriginal'])

ldap_org_4['OrgCodeLevel4'] =  pd.to_numeric(ldap_org_4['OrgCodeLevel4'])
ldap_org_4['ParentCodeLevel3'] =  pd.to_numeric(ldap_org_4['ParentCodeLevel3'])

ldap_org_4 = ldap_org_4[['OrgCodeLevel4','OrgNameLevel4','Level3Reference','ParentCodeLevel3','createTimeStamp','modifyTimeStamp','ManagerEMPLID']]

org_4_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\BU\Target Files\LDAP_BU_4" + '.xlsx', engine='xlsxwriter')
ldap_org_4.to_excel(org_4_writer, 'cads_bu_merged', index=False)
org_4_writer.save()

'''Org Code Level 5'''

ldap_org_5 = ldap_cads_bu.copy()

ldap_org_5 = ldap_org_5[((ldap_org_5['Organization'].str[12:18] == '000000') & (ldap_org_5['Organization'].str[10:12] > '00'))
 | ((ldap_org_5['Organization'].str[10:18] == '00000001') & (ldap_org_5['Organization'].str[6:10] > '0000'))]  

ldap_org_5['Level4Reference'] = ""
ldap_org_5['Level4Reference'] = ldap_org_5.apply(lambda x: ldap_org_5['Parent'])

ldap_org_5['ParentCodeLevel4'] = ''
ldap_org_5['ParentCodeLevel4'] = ldap_org_5.apply(lambda x: ldap_org_5['Organization'].str[:10] + '00000000')

ldap_org_5['OrgCodeLevel5'] = ""
ldap_org_5['OrgNameLevel5'] = ""
ldap_org_5['OrgCodeLevel5'] = ldap_org_5.apply(lambda x: ldap_org_5['Organization'])
ldap_org_5['OrgNameLevel5'] = ldap_org_5.apply(lambda x: ldap_org_5['OrgNameOriginal'])

ldap_org_5['OrgCodeLevel5'] =  pd.to_numeric(ldap_org_5['OrgCodeLevel5'])
ldap_org_5['ParentCodeLevel4'] =  pd.to_numeric(ldap_org_5['ParentCodeLevel4'])

ldap_org_5 = ldap_org_5[['OrgCodeLevel5','OrgNameLevel5','Level4Reference','ParentCodeLevel4','createTimeStamp','modifyTimeStamp','ManagerEMPLID']]


org_5_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\BU\Target Files\LDAP_BU_5" + '.xlsx', engine='xlsxwriter')
ldap_org_5.to_excel(org_5_writer, 'cads_bu_merged', index=False)
org_5_writer.save()


'''Org Code Level 6'''

ldap_org_6 = ldap_cads_bu.copy()

ldap_org_6 = ldap_org_6[((ldap_org_6['Organization'].str[14:18] == '0000') & (ldap_org_6['Organization'].str[12:14] > '00'))
 | ((ldap_org_6['Organization'].str[12:18] == '000001') & (ldap_org_6['Organization'].str[10:12] > '00'))]  


ldap_org_6['Level5Reference'] = ""
ldap_org_6['Level5Reference'] = ldap_org_6.apply(lambda x: ldap_org_6['Parent'])

ldap_org_6['ParentCodeLevel5'] = ''
ldap_org_6['ParentCodeLevel5'] = ldap_org_6.apply(lambda x: ldap_org_6['Organization'].str[:12] + '000000')

ldap_org_6['OrgCodeLevel6'] = ""
ldap_org_6['OrgNameLevel6'] = ""
ldap_org_6['OrgCodeLevel6'] = ldap_org_6.apply(lambda x: ldap_org_6['Organization'])
ldap_org_6['OrgNameLevel6'] = ldap_org_6.apply(lambda x: ldap_org_6['OrgNameOriginal'])

ldap_org_6['OrgCodeLevel6'] =  pd.to_numeric(ldap_org_6['OrgCodeLevel6'])
ldap_org_6['ParentCodeLevel5'] =  pd.to_numeric(ldap_org_6['ParentCodeLevel5'])

ldap_org_6 = ldap_org_6[['OrgCodeLevel6','OrgNameLevel6','Level5Reference','ParentCodeLevel5','createTimeStamp','modifyTimeStamp','ManagerEMPLID']]


org_6_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\BU\Target Files\LDAP_BU_6" + '.xlsx', engine='xlsxwriter')
ldap_org_6.to_excel(org_6_writer, 'cads_bu_merged', index=False)
org_6_writer.save()

'''Org Code Level 7'''

ldap_org_7 = ldap_cads_bu.copy()

ldap_org_7 = ldap_org_7[(ldap_org_7['Organization'].str[16:18] == '00') & (ldap_org_7['Organization'].str[14:16] > '00')]  


ldap_org_7['Level6Reference'] = ""
ldap_org_7['Level6Reference'] = ldap_org_7.apply(lambda x: ldap_org_7['Parent'])

ldap_org_7['ParentCodeLevel6'] = ''
ldap_org_7['ParentCodeLevel6'] = ldap_org_7.apply(lambda x: ldap_org_7['Organization'].str[:14] + '0000')

ldap_org_7['OrgCodeLevel7'] = ""
ldap_org_7['OrgNameLevel7'] = ""
ldap_org_7['OrgCodeLevel7'] = ldap_org_7.apply(lambda x: ldap_org_7['Organization'])
ldap_org_7['OrgNameLevel7'] = ldap_org_7.apply(lambda x: ldap_org_7['OrgNameOriginal'])

ldap_org_7['OrgCodeLevel7'] =  pd.to_numeric(ldap_org_7['OrgCodeLevel7'])
ldap_org_7['ParentCodeLevel6'] =  pd.to_numeric(ldap_org_7['ParentCodeLevel6'])

ldap_org_7 = ldap_org_7[['OrgCodeLevel7','OrgNameLevel7','Level6Reference','ParentCodeLevel6','createTimeStamp','modifyTimeStamp','ManagerEMPLID']]


org_7_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\BU\Target Files\LDAP_BU_7" + '.xlsx', engine='xlsxwriter')
ldap_org_7.to_excel(org_7_writer, 'cads_bu_merged', index=False)
org_7_writer.save()

'''Org Code Level 8'''

ldap_org_8 = ldap_cads_bu.copy()

ldap_org_8 = ldap_org_8[(ldap_org_8['Organization'].str[14:16] > '00') & (ldap_org_8['Organization'].str[16:18] > '00')]  


ldap_org_8['Level7Reference'] = ""
ldap_org_8['Level7Reference'] = ldap_org_8.apply(lambda x: ldap_org_8['Parent'])

ldap_org_8['ParentCodeLevel7'] = ''
ldap_org_8['ParentCodeLevel7'] = ldap_org_8.apply(lambda x: ldap_org_8['Organization'].str[:16] + '00')

ldap_org_8['OrgCodeLevel8'] = ""
ldap_org_8['OrgNameLevel8'] = ""
ldap_org_8['OrgCodeLevel8'] = ldap_org_8.apply(lambda x: ldap_org_8['Organization'])
ldap_org_8['OrgNameLevel8'] = ldap_org_8.apply(lambda x: ldap_org_8['OrgNameOriginal'])

ldap_org_8['OrgCodeLevel8'] =  pd.to_numeric(ldap_org_8['OrgCodeLevel8'])
ldap_org_8['ParentCodeLevel7'] =  pd.to_numeric(ldap_org_8['ParentCodeLevel7'])

ldap_org_8 = ldap_org_8[['OrgCodeLevel8','OrgNameLevel8','Level7Reference','ParentCodeLevel7','createTimeStamp','modifyTimeStamp','ManagerEMPLID']]


org_8_writer = pd.ExcelWriter(r"C:\Users\eakarsu\Documents\Python\CADS\BU\Target Files\LDAP_BU_8" + '.xlsx', engine='xlsxwriter')
ldap_org_8.to_excel(org_8_writer, 'cads_bu_merged', index=False)
org_8_writer.save()


