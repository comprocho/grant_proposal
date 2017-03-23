import numpy as np
import pandas as pd
import re

#Imports excel files
dfa = pd.read_excel('A.Summary Charges FY17.xlsx', skip_footer=1)
dfb = pd.read_excel('B.List of PIs.xlsx')

dfc_bc = pd.read_excel('C.membership report.xls', sheetname=0, skiprows=1, skip_footer = 1, parse_cols=[0,1])
dfc_cpc = pd.read_excel('C.membership report.xls', sheetname=1, skiprows=1, skip_footer = 1, parse_cols=[0,1])
dfc_et = pd.read_excel('C.membership report.xls', sheetname=2, skiprows=1, skip_footer = 1, parse_cols=[0,1])
dfc_mo = pd.read_excel('C.membership report.xls', sheetname=3, skiprows=1, skip_footer = 1, parse_cols=[0,1])

#Combines dfc sub-dataframes to the dfc
frames = [dfc_bc, dfc_cpc, dfc_et, dfc_mo]
dfc = pd.concat(frames).reset_index()

dfc = dfc[['Program', 'Member Name']]
dfc['Member Name'] = dfc['Member Name'].str.split(', ') #gets a last name of PI

dfc['LN'] = dfc['Member Name'].str.get(0) #Slices the list and create last name column 
dfc['FN'] = dfc['Member Name'].str.get(1) #Slices the list and create last name column 
dfc = dfc[['Program', 'LN', 'FN']]

#Drops duplicated items and inserts first character of first name to the end of last name
dfc.loc[dfc.LN.duplicated(keep=False), 'LN'] += ' ' + dfc.FN.str[0]

#Older version
# df1 = df.loc[df.duplicated('LN', False)]
# df2 = pd.DataFrame(df1.LN + ' '+ df1.FN.str.get(0))
# df3 = pd.concat([df1,df2], axis=1)
# df3 = df3[[0, 'FN']]
# df3.columns = ['LN', 'FN']
# df.update(df3)

dfa = dfa.PI.drop_duplicates().dropna().to_frame().reset_index() 
dfa = dfa[['PI']]

dfb = dfb[['PI Last Name']]
dfb = dfb.rename(columns={'PI Last Name': 'PI'})
dfb.PI = dfb.PI.replace({'Shajahan': 'Shajahan-Haq'})

frames = [dfa, dfb]
df_ab = pd.concat(frames) #Concatenates dfa and dfb, creates a list of PI who have used LCCC

df_ab = df_ab[['PI']]
df_ab.PI = df_ab.PI.replace({'Wang':'Wang D', 
                             'Bouker/Clarke': 'Clarke', 
                             'deAssis': 'De Assis',
                             'Johnson': 'Johnson M', 
                             'Chang': 'Chang E', 
                             'Hilakivi Clarke': 'Hilakivi-Clarke', 
                             'Pohlman': 'Pohlmann'})
df_ab = df_ab.drop_duplicates().sort_values(by='PI').reset_index()
df_ab = df_ab[['PI']]
df_ab['Program'] = pd.Series().fillna(' ')
# df_ab['Program'] = df_ab['Program'].fillna(' ')

#Updates Program column on df_ab with the program name
x1 = df_ab.set_index(['PI'])['Program']
x2 = dfc.set_index(['LN'])['Program']
x1.update(x2)
df_ab['Program'] = x1.values

df_ab = df_ab[['Program', 'PI']] #Changes the order of columns
df_ab.Program = df_ab.Program.fillna('Other User')
df_ab = df_ab.sort_values(by='Program')

df_ab.to_excel('List of All LCCC PI.xlsx')

total_num_used_sr = len(df_ab) #Total number of lccc members

other_user = df_ab[df_ab.Program == 'Other User'] #Gets LCCC PI who does not belong to any four programs
num_others = len(other_user) #Number of SR users who do not belong to any four programs

def peer_reviewed():
    peer_reviewed = pd.read_excel('D.CurrentPeer-review funded Investigators by Program.xlsx', skiprows = 1)

    peer_reviewed.MO = peer_reviewed.MO.str.replace(r'\d*\)', '').str.lstrip().str.rstrip()
    peer_reviewed['Program'] = pd.Series()

    peer_reviewed.iloc[0:10, peer_reviewed.columns.get_loc('Program')] = 'MO'
    peer_reviewed.iloc[12:21, peer_reviewed.columns.get_loc('Program')] = 'BC'
    peer_reviewed.iloc[22:32, peer_reviewed.columns.get_loc('Program')] = 'ET'
    peer_reviewed.iloc[33:48, peer_reviewed.columns.get_loc('Program')] = 'CPC'

    peer_reviewed = peer_reviewed.drop(peer_reviewed.index[[10,11,21,32]])
    peer_reviewed.columns = ['PI', 'Program']

    peer_reviewed = peer_reviewed[['Program', 'PI']]

    peer_reviewed.PI = peer_reviewed.PI.str.replace('Avantagiatti', 'Avantaggiati')

    peer_reviewed.PI = peer_reviewed.PI.str.split(' ').str.get(-1).str.replace('Wang', 'Wang J').str.replace('Assis', 'De Assis')
    
    return peer_reviewed
peer_reviewed()
total_num_peer_reviewed= len(peer_reviewed())

peer_reviewed().to_excel('List of All Peer-Review PI by Program.xlsx')

df_lccc_peer_reviewed = df_ab[df_ab['PI'].isin(peer_reviewed()['PI'])] #filters a List of PI who belongs to peer-reviwed funding group

df_lccc_peer_reviewed.to_excel('List of Peer-Reviewed LCCC PI.xlsx')

num_lccc_peer_reviewed =len(df_lccc_peer_reviewed) #gets a number of LCCC PIs who are members of peer-reviwed funding group

other_user_dropped = df_ab[~(df_ab['Program'] == 'Other User')]
##drops a list of PI with no program assigned and contains only those who belong to any four programs
#other_pi_dropped.replace(' ', np.nan, inplace = True)
#other_pi_dropped.dropna(subset=['Program'], inplace=True)

non_peer_reviewed = other_user_dropped[~(other_user_dropped['PI'].isin(peer_reviewed()['PI']))]
non_peer_reviewed.to_excel('List of Non Peer-Review LCCC PI.xlsx')

num_lccc_non_peer_reviewed = len(non_peer_reviewed)
num_sr_users_other_pi_dropped= len(other_user_dropped) #gets a number of PIs who do not belong to any of four programs

index = ['LCCC members, peer-reviewed funding', 'LCCC members, non–peer-reviewed funding','Other users','Total']
columns = ['No. Users', '% Users']
table_five = pd.DataFrame(index=index, columns = columns)

table_five.iloc[0,0] = num_lccc_peer_reviewed
table_five.iloc[1,0] = num_lccc_non_peer_reviewed
table_five.iloc[2,0] = num_others
table_five.iloc[3,0] = table_five['No. Users'].sum()

table_five['% Users'] = (table_five['No. Users'] / table_five['No. Users'].loc['Total']) * 100
table_five['% Users'] = table_five['% Users'].astype(float).round(1)

table_five.to_excel('table_five.xlsx')
table_five

index = ['BC', 'CPC', 'ET', 'MO', 'Total']
columns = ['Peer Reviewed', 'Non-Peer Reviewed']
table_two = pd.DataFrame(index=index, columns = columns)
table_two.index.names = ['Program']
table_two = table_two.fillna('0')

table_two['Peer Reviewed'] = df_lccc_peer_reviewed.groupby('Program').size()
table_two['Non-Peer Reviewed'] = non_peer_reviewed.groupby('Program').size()

table_two.loc['Total', 'Peer Reviewed'] = table_two['Peer Reviewed'].sum()
table_two.loc['Total', 'Non-Peer Reviewed'] = table_two['Non-Peer Reviewed'].sum()

table_two = table_two.astype(int)
table_two

table_two.to_excel('table_two.xlsx')