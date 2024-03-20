# coding: utf-8
import pandas as pd
pd.options.mode.chained_assignment = None

def GetLidnummer(invoer,column_naam='Naam'):
    # krijg lidnummer gebaseerd op naam en voornaam
    # input: 
    #   invoer: gpd met naam en voornaam
    #   column_naam: kolom met naam in

    # get export
    export,clubs=getExport()
    export['naam_combi']=(export['Naam']+' '+export['Voornaam']).str.lower().str.strip()

    # lower
    invoer['naam_combi']=invoer[column_naam].str.lower().str.strip()

    # combineer voor personen
    res=pd.merge(invoer,export[['naam_combi','Lidnummer','Club (0)']],how='left',left_on='naam_combi',right_on='naam_combi')
    print('Issues: ',res[res['Lidnummer'].isnull()])
    noissues=res[res['Lidnummer'].notnull()]

    # return 
    return noissues[[column_naam,'Lidnummer','Club (0)']]


def GetMailinglijst_Naam(invoer,column_naam='Naam',functies=['Secretaris','Voorzitter']):
    # krijg mailinglijst gebaseerd op naam en voornaam
    # input: 
    #   invoer: gpd met naam en voornaam
    #   column_naam: kolom met naam in
    #   column_voornaam:  kolom met voornaam in 
    #   functies: functies om te mailen
            
    
    # get export
    export,clubs=getExport()
    export['naam_combi']=(export['Naam']+' '+export['Voornaam']).str.lower().str.strip()
    
    # lower
    invoer['naam_combi']=invoer[column_naam].str.lower().str.strip()

    
    # combineer voor personen
    res=pd.merge(invoer,export[['naam_combi','Email', 'Email (CC)','Club (0)','Lidnummer']],how='left',left_on='naam_combi',right_on='naam_combi')
    print('Issues because not found: ',res[res['Club (0)'].isnull()])
    noissues=res[res['Club (0)'].notnull()]
    sel1=noissues[noissues['Email'].notnull()]['Email']
    sel2=noissues[noissues['Email (CC)'].notnull()]['Email (CC)']
    lijst=','.join(sel1.astype(str))+','+(','.join(sel2.astype(str)))
    
    # lijst van clubadressen
    clublijst=noissues['Club (0)'].unique()
    lijst_clubs=GetMailClubs(clublijst,functies=functies)   
        
    # vewerken
    print('Geen mailadres\n',noissues[noissues['Email'].isnull()][['Club (0)','naam_combi','Lidnummer']])
    print('\n')

    print(lijst+','+lijst_clubs)

    
def GetMailClubs(clublist,functies=['secretaris','voorzitter']):
    # krijg mailinglijst van clubs gebaseerd op de functies en 
    # invoer:
    #   clublist: lijst met clubs
    #   functies: lijst met functies
    
    # check options
    options_functies=['interclubJeugd',
         'interclubSeniors',
         'secretaris',
         'voorzitter',
         'penningmeester',
         'aanspreekpunt']

    # Check if all elements in 'functies' are part of 'options_functies'
    not_found = [functie for functie in functies if functie not in options_functies]

    # Provide error message if there are elements not found in 'options_functies'
    if len(not_found)>0:
        error_message = f"The following roles are not correct: {', '.join(not_found)}"   
        print(error_message)
        print('Exiting')
        import sys
        sys.exit()
        
    # get export
    export,clubs=getExport()
    
    # selectie kolommen
    columns_to_retrieve = [f"{functie} Emails" for functie in functies]
    retrieved_columns_df = clubs[columns_to_retrieve]
    # get mails
    mails_clubs=clubs[clubs.index.isin(clublist)][columns_to_retrieve].fillna('hans.hooyberghs@gmail.com')
    lijst_clubs=','.join(mails_clubs.values.flatten())    
    
    # get unique values
    unique_emails = ','.join(set(email.strip() for email in lijst_clubs.split(',') if email.strip()))

    return unique_emails





def GetMailinglijst_Lidnummer(invoer,column_lidnummer='Lidnummer',functies=['secretaris','voorzitter']):
    # krijg mailinglijst gebaseerd op lidnummer
    # input: 
    #   invoer: gpd met naam en voornaam
    #   column_lidnummer: kolom met lidnummer in
    #   functies: functies om te mailen
    
    # get export
    export,clubs=getExport()
    

    
    # combineer voor personen
    res=pd.merge(invoer,export[['Lidnummer','Email', 'Email (CC)','Club (0)','Naam','Voornaam']],how='left',left_on=column_lidnummer,right_on='Lidnummer')
    print('Issues: ',res[res['Club (0)'].isnull()])
    noissues=res[res['Club (0)'].notnull()]
    sel1=noissues[noissues['Email'].notnull()]['Email']
    sel2=noissues[noissues['Email (CC)'].notnull()]['Email (CC)']
    lijst=','.join(sel1.astype(str))+','+(','.join(sel2.astype(str)))
    unique_emails = ','.join(set(email.strip() for email in lijst.split(',') if email.strip()))

    
    # lijst van clubadressen
    clublijst=noissues['Club (0)'].unique()
    lijst_clubs=GetMailClubs(clublijst,functies=functies)   
    
    # vewerken
    print('Geen mailadres\n',noissues[noissues['Email'].isnull()][['Club (0)','Naam','Voornaam']])
    print('\n')
    
    print(unique_emails+','+lijst_clubs)


def getExport():
    import os
    # lees export in
    if os.getenv('EXPORT_TT'):
        export_tt = os.getenv('EXPORT_TT')
    else:
        export_tt = str(input(f"Environmental variable linking to database not set. Please enter the location manually."))

    export=pd.read_csv(export_tt,sep=';',encoding='unicode_escape')
    export['naam_combi']=(export['Naam']+ ' '+export['Voornaam']).str.lower().str.strip() 

    # get club information
    clubs=GetClubs(export)

    # add names of clubs
    clubdata = {
    'Club (0)': ["A003", "A008", "A062", "A074", "A075", "A095", "A097", "A105", "A115", "A117", "A118", "A123", 
                 "A127", "A129", "A130", "A135", "A136", "A138", "A139", "A141", "A142", "A147", "A155", "A159", 
                 "A160", "A167", "A176", "A182", "A186", "A201", "A211", "A212", "A218", "A222"],
    'NaamClub': ["KTTC Salamander Mechelen", "TTC Brasgata", "KTTC  A.F.P. Antwerpen", "KTTC Sporting Hove", 
                 "TTC Rupel", "TTK Turnhout", "KTTC Nijlen-Bevel", "TTK Dessel", "TTK Buhlmann Berlaar", "Geelse TTC", 
                 "TTK Rijkevorsel", "TTC Virtus", "TTC Retie", "Borsbeek TTK", "TTK Schoten", "TTK Minderhout", 
                 "TTC Zoersel", "TTC Blue Rackets", "KTTC Hallaar", "TTK Merksplas", "TTC Tecemo", "TTK Gierle", 
                 "TTC Wommelgem", "TTK Real", "TTC Walem", "TTC Lille \"Boemerangske\"", "TTC Sokah Hoboken", 
                 "TTC Nodo", "TTC Willebroek", "TTC Hulshout", "KTTK Zwijndrecht", "TTC Sint Antonius", "TTV Poppel", 
                 "Omni Vrembo"]
    }
    clubdata=pd.DataFrame(clubdata)
    export=export.merge(clubdata,on='Club (0)')   
    clubs=clubs.merge(clubdata,left_index=True,right_on='Club (0)')
    clubs.index=clubs['Club (0)']
    return export,clubs


def GetClubs(df):
    # Identify columns that start with "Club fun" (case insensitive)
    club_fun_columns = [col for col in df.columns if col.lower().startswith("club fun")]

    # Filter rows where at least one of the columns that start with "Club fun" is not empty
    filtered_df = df.dropna(subset=club_fun_columns, how='all')

    # Sort the DataFrame based on the "Club [0]" column, if available
    if "Club [0]" in df.columns:
        filtered_df = filtered_df.sort_values(by="Club [0]")

    # Convert float NaNs to empty strings in the 'Club fun' columns
    filtered_df[club_fun_columns] = filtered_df[club_fun_columns].fillna('')

    # Add a combined "Club fun" column
    filtered_df['Club fun Combined'] = filtered_df[club_fun_columns].agg(','.join, axis=1)

    # Retain all columns, including the new "Club fun Combined"
    final_df = filtered_df.assign(**{'Club fun Combined': filtered_df['Club fun Combined']})

    # Initialize a new DataFrame to store the reshaped data
    reshaped_df_with_combined_emails = pd.DataFrame()

    # Create a new column combining 'Naam' and 'Voornaam'
    final_df['Full Name'] = final_df['Naam'] + ' ' + final_df['Voornaam']
    final_df['Combined Email'] = final_df['Email'].fillna('') + final_df['Email (CC)'].apply(lambda x: ',' + x if pd.notna(x) else '')

    # List of roles we are interested in
    roles_of_interest = ['interclubJeugd', 'interclubSeniors', 'secretaris', 'voorzitter', 'penningmeester', 'aanspreekpunt']

    # Loop through each role to create the respective columns
    for role in roles_of_interest:
        # Find rows where the role appears in 'Club fun Combined'
        role_df = final_df[final_df['Club fun Combined'].str.contains(role, case=False, na=False)]
        
        # Group by 'Club (0)' and aggregate full names and combined emails, then convert lists to comma-separated strings
        role_group_names = role_df.groupby('Club (0)')['Full Name'].apply(lambda x: ', '.join(x)).reset_index()
        role_group_emails = role_df.groupby('Club (0)')['Combined Email'].apply(lambda x: ', '.join(x)).reset_index()
        
        # Rename columns and merge with reshaped_df_with_combined_emails
        role_group_names.columns = ['Club (0)', f"{role} Names"]
        role_group_emails.columns = ['Club (0)', f"{role} Emails"]
        
        if reshaped_df_with_combined_emails.empty:
            reshaped_df_with_combined_emails = role_group_names
        else:
            reshaped_df_with_combined_emails = pd.merge(reshaped_df_with_combined_emails, role_group_names, on='Club (0)', how='outer')
        
        reshaped_df_with_combined_emails = pd.merge(reshaped_df_with_combined_emails, role_group_emails, on='Club (0)', how='outer')


    reshaped_df_with_combined_emails.set_index('Club (0)', inplace=True)
    
    return reshaped_df_with_combined_emails

    

