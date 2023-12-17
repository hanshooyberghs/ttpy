##############
import zeep
from zeep import helpers
from zeep.wsse.username import UsernameToken
import pandas as pd
from ttpy import mailroutines
from ttpy.generalroutines import Replace,AddTable,SaveExcel
from docx import Document
import os


def GetTournamentEntries(tornooien,inschrijvingsgeld,provincie='A'):
    # Get a list of tournaments
    # input:
    #   tornooien: lijst van tornooien (zoals in competitiesite)
    #   inschrijvingsgeld: inschrijvingsgeld
    #   provincie: selectie op provincie  (default: Antwerpen)
    
    # client aanmaken
    wsdl='http://api.vttl.be/0.7/?wsdl'
    client = zeep.Client(wsdl=wsdl)

    # tornooien inladen
    tournaments=client.service.GetTournaments()
    input_dict = helpers.serialize_object(tournaments)


    # find tournaments and players
    print('Lees tornooien')
    dfs=[]
    for item in input_dict['TournamentEntries']:
        if item['Name'] in tornooien:
            print(item['Name'],' ',item['UniqueIndex'])
            spelers_tornooi=client.service.GetTournaments(WithRegistrations=True,TournamentUniqueIndex=item['UniqueIndex'])
            spelers_tornooi = helpers.serialize_object(spelers_tornooi)
            spelers=spelers_tornooi['TournamentEntries'][0]['SerieEntries']
            
            # loop over all series
            for serie in spelers:       
                nummers=[w['Member']['UniqueIndex'] for w in serie['RegistrationEntries']]
                nummers=pd.DataFrame(nummers,columns=['Lidnummer'])
                nummers['Tornooi']=item['Name'] 
                nummers['reeks']=item['Name']+' '+serie['Name']
                nummers['Naam']=[w['Member']['LastName'] for w in serie['RegistrationEntries']]
                nummers['Voornaam']=[w['Member']['FirstName'] for w in serie['RegistrationEntries']]
                nummers['Club']=[w['Club']['UniqueIndex'] for w in serie['RegistrationEntries']]
                nummers['Inschrijvingsgeld']=inschrijvingsgeld
                dfs.append(nummers)


    #combine all
    all_registrations=pd.concat(dfs)

    # only one participation in the same tournament
    print('Nog aanpassen voor dubbels: wel tweemaal in zelfde tornooi, tornooi anders noemen')
    all_registrations=all_registrations.drop_duplicates(['Lidnummer','Naam','Voornaam','Club','Tornooi'])

    # select province
    all_registrations=all_registrations[[w[0]=='A' for w in all_registrations.Club]]

    # convert to simple list
    combi_amount=all_registrations.groupby(['Lidnummer','Naam','Voornaam','Club'])[['Inschrijvingsgeld']].sum()
    combi_list_tournaments=all_registrations.groupby(['Lidnummer','Naam','Voornaam','Club'])['Tornooi'].apply(list).apply(lambda p:','.join(p))
    final=combi_amount.merge(combi_list_tournaments,left_index=True,right_index=True).reset_index()
    
    return final,all_registrations

def AddFines(final,file_boetes,tag_lidnummer='Lidnummer',tag_supplement='Supplement'):
    # add fines (non participation), merging on column Lidnummer
    #  input:
    #   final: input dataframe with entries
    #   boets: excel file with fines. Should contain columns Lidnummer, Supplement and Reden Supplement
    
    # load input and check
    boetes=pd.read_excel(file_boetes)
        
    # checks
    missing_columns = [column for column in [tag_lidnummer,tag_supplement] if column not in boetes.columns]
    
    if missing_columns:
        raise ValueError(f"Missing columns in DataFrame: {', '.join(missing_columns)}")
        sys.exit(1)

    # merge
    final=final.merge(boetes.drop(columns='Naam'),on='Lidnummer',how='left')
    final['Totaal']=final['Inschrijvingsgeld']+final['Supplement'].fillna(0)
    
    return final
    
def PrintTotals(final,all_registrations,uitvoer_detail_totaal,column_supplement='Supplement',column_reason_supplement='Reden Supplement'):
    ## Print totalen ##
    # input:
    #   final: final dataframe from load routines
    #   all_registrations: list with all registrations, not grouped, from load routines
    #   uitvoer_detail_totaal: file in which to save the data
    ###################
    
    print('\nTotalen')
    print(all_registrations.groupby('Tornooi').sum()['Inschrijvingsgeld'])
    
    print('\nTotale supplementen')
    print(final.groupby(column_reason_supplement).sum()[column_supplement])

    # opslaan lijst nog te betalen
    df=final.sort_values('Naam')[['Lidnummer', 'Naam','Voornaam','Club', 'Totaal','Inschrijvingsgeld', column_supplement,
          'Tornooi',column_reason_supplement]].rename(columns={'Totaal':'Totaal (euro)'})
    SaveExcel(df,uitvoer_detail_totaal,'Saldo')


def TournamentsSaveAndMail(final,clubs_direct,default_items,uitvoer_detail_club,uitvoer_detail_individueel,lege_factuur,tornooien,startnummer_factuur=0,formaat_factuur='A-2023/',column_supplement='Supplement',column_reason_supplement='Reden Supplement'):
    ## Make MailingLists and save data
    # input:
    #   final: final dataframe from load routines
    #   clubs_direct: clubs with direct payment
    #   default_items: default items to add to the word document
    #   uitvoer_detail_individueel: savefile for individual data
    #   uitvoer_detail_club: savefile for clubdata
    ###################
    
    
    # selectie clubs die niet rechtsteeks betalen
    #############################################
    print('\nAfrekening per speler')
    selectie_betalen=final[~final.Club.isin(clubs_direct)]

    # mailadressen
    mailroutines.GetMailinglijst_Lidnummer(selectie_betalen[['Lidnummer']],functies=['secretaris','voorzitter','penningmeester'])

    # opslaan lijst in Excel
    df=selectie_betalen.sort_values('Naam')[['Lidnummer', 'Naam','Voornaam','Club', 'Totaal','Inschrijvingsgeld', column_supplement,
          'Tornooi',column_reason_supplement]].rename(columns={'Totaal':'Totaal (euro)'})
    SaveExcel(df,uitvoer_detail_individueel,'Saldo')

    # verwerking clubs die rechtsteeks betalen
    #############################################
    print('\nAfrekening per club, direct betaling')
    selectie_betalen_club=final[final.Club.isin(clubs_direct)]

    # mailadressen
    mailclubs=mailroutines.GetMailClubs(clubs_direct,functies=['secretaris','voorzitter','penningmeester'])
    print(mailclubs)
    
    # opslaan in lijst met excel
    df=selectie_betalen_club.sort_values(['Club','Naam'])[['Lidnummer', 'Naam','Voornaam','Club', 'Totaal','Inschrijvingsgeld', column_supplement,
          'Tornooi',column_reason_supplement]]
    SaveExcel(df,uitvoer_detail_club,'Saldo') 
    
    # add totals
    with pd.ExcelWriter(uitvoer_detail_club, mode='a', engine='openpyxl') as writer:
        selectie_betalen_club.groupby('Club').sum()['Totaal'].to_excel(writer,sheet_name='Totalen')

    
    # make word files
    data=selectie_betalen_club.copy()
    cnt=0
    dfs=[]
    for clubnummer in clubs_direct:
        doc = Document(lege_factuur)
        sel=data[data.Club==clubnummer]
        
        if len(sel)>0:
            # prices for clubs
            totaal=sel['Totaal'].sum()
            
            # add to replace
            replace=default_items.copy()
            replace['CLUBNUMMER']=clubnummer
            replace['LIJST']=', '.join(tornooien)
            replace['TOTAAL']=str(totaal)
            
            factuurnummer=formaat_factuur+str(startnummer_factuur+cnt).zfill(3)
            replace['NUMMER']=factuurnummer

            
            # make string replacements
            doc=Replace(replace,doc)
            
            # add table
            tmp=sel[['Naam','Voornaam','Totaal','Inschrijvingsgeld',column_supplement,column_reason_supplement]].fillna('').sort_values('Naam')
            doc=AddTable('TABEL',tmp,doc)

            
            doc.save(os.path.dirname(uitvoer_detail_club)+'/'+clubnummer+'.docx')
            
            cnt=cnt+1

def Inschrijving(lidnummer,tornooi,reeks=None,mail=True,dubbel=False):
    # Inschrijven voor tornooi
    # input:
    #   lidnummer: lidnummer (of lijst van lidnummers bij dubbel)
    #   tornooi: tornooi
    #   reeks: reeks (optioneel, indien niet gegeven: automatisch bepaald)
    #   mail: send mail or not
    #   dubbel: gaat het om een dubbel of niet
    ###################
    
    # client aanmaken
    wsdl='http://api.vttl.be/0.7/?wsdl'
    client = zeep.Client(wsdl=wsdl)
    # tornooien inladen
    tournaments=client.service.GetTournaments()
    input_dict = helpers.serialize_object(tournaments)

    # lees paswoord van environmental variables
    if os.getenv('ACCOUNT_TT'):
        account = os.getenv('ACCOUNT_TT')
    else:
        account = str(input(f"Environmental variable linking to account not set. Please enter the location manually."))
    if os.getenv('PASWOORD_TT'):
        paswoord = os.getenv('PASWOORD_TT')
    else:
        paswoord = str(input(f"Environmental variable linking to password not set. Please enter the location manually."))

    # achterhaal codes
    for item in input_dict['TournamentEntries']:
        if item['Name'].lower()==tornooi.lower():
            for item2 in item['SerieEntries']:
                if item2['Name'].lower()==reeks.lower():
                    index_tornooi=item['UniqueIndex']
                    index_reeks=item2['UniqueIndex']

    # check whether index_tornooi exists
    if 'index_tornooi' not in locals():
        print('Tornooi of reeks niet gevonden, controleer input')
        import sys
        sys.exit()

    # inschrijven   
    try:
        if dubbel:
            # check whether lidnummer is array
            if isinstance(lidnummer, list):
                response=client.service.TournamentRegister(TournamentUniqueIndex=index_tornooi,SerieUniqueIndex=index_reeks,PlayerUniqueIndex=lidnummer,Credentials={'Account':account,'Password':paswoord},NotifyPlayer=mail)
            else:
                raise ValueError(f"Inschrijving dubbel: lidnummer is geen lijst")

        else:
            # check whether lidnummer is integer 
            if isinstance(lidnummer, int):
                response=client.service.TournamentRegister(TournamentUniqueIndex=index_tornooi,SerieUniqueIndex=index_reeks,PlayerUniqueIndex=str(lidnummer),Credentials={'Account':account,'Password':paswoord},NotifyPlayer=mail)
            else:
                raise ValueError(f"Inschrijving enkel: lijsten van lidnummers gegeven")

        print('Inschrijving voor ',lidnummer,' in ',tornooi,' in reeks ',reeks,' is gelukt.')
    except ValueError as ve:
        print(f"Caught an error: {ve}")               
    except zeep.exceptions.Fault as fault:
        print(f"Caught a zeep Fault: {fault}")
  

