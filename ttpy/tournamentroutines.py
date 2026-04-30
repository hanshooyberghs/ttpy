"""
tournamentroutines.py
---------------------
Routines voor het ophalen, verwerken en doorsturen van tornooi-inschrijvingen
via de VTTL SOAP-API (zeep).

Functies:
    GetTournamentEntries()    – Haal inschrijvingen op voor een lijst tornooien.
    AddFines()                – Voeg boetes (niet-deelname) toe aan de inschrijvingslijst.
    PrintTotals()             – Druk totalen af en sla detail op als Excel.
    TournamentsSaveAndMail()  – Sla Excel- en Word-facturen op en verstuur e-mails.
    Inschrijving()            – Schrijf een speler in via lidnummer.
    InschrijvingNaam()        – Schrijf een speler in via naam.

Vereiste omgevingsvariabelen:
    ACCOUNT_TT   – VTTL API-accountnaam.
    PASWOORD_TT  – VTTL API-wachtwoord.
"""
##############
import sys
import os
import zeep
from zeep import helpers
from zeep.wsse.username import UsernameToken
import pandas as pd
from ttpy import mailroutines
from ttpy.generalroutines import Replace, AddTable, SaveExcel
from docx import Document


WSDL = 'http://api.vttl.be/0.7/?wsdl'


def _create_client():
    """Maak een zeep SOAP-client aan en haal de tornooidatalijst op.

    Returns:
        tuple[zeep.Client, dict]: De SOAP-client en een geserialiseerd
        woordenboek met alle tornooigegevens.
    """
    tournaments = client.service.GetTournaments()
    input_dict = helpers.serialize_object(tournaments)
    return client, input_dict


def _get_credentials():
    """Haal VTTL API-inloggegevens op uit omgevingsvariabelen of gebruikersinvoer.

    Returns:
        tuple[str, str]: Account en wachtwoord voor de VTTL API.
    """
    paswoord = os.getenv('PASWOORD_TT') or input("Environmental variable PASWOORD_TT not set. Please enter manually: ")
    return account, paswoord


def _lookup_tournament_indices(input_dict, tornooi, reeks):
    """Zoek de unieke indices op van een tornooi en een reeks.

    Args:
        input_dict (dict): Geserialiseerde tornooilijst afkomstig van de VTTL API.
        tornooi (str): Naam van het tornooi (hoofdletterongevoelig).
        reeks (str): Naam van de reeks binnen het tornooi (hoofdletterongevoelig).

    Returns:
        tuple[int, int]: ``(tornooi_index, reeks_index)`` als unieke API-indices.

    Raises:
        SystemExit: Als ``reeks`` niet opgegeven is, of als het tornooi of de
            reeks niet gevonden wordt in de API-data.
    """
    if reeks is None:
        print('Reeks moet opgegeven worden')
        sys.exit()
    for item in input_dict['TournamentEntries']:
        if item['Name'].lower() == tornooi.lower():
            for item2 in item['SerieEntries']:
                if item2['Name'].lower() == reeks.lower():
                    return item['UniqueIndex'], item2['UniqueIndex']
    print('Tornooi of reeks niet gevonden, controleer input')
    print(tornooi, reeks)
    sys.exit()

def GetTournamentEntries(tornooien, inschrijvingsgeld, file_dubbels='Dubbels.xlsx', provincie='A', dubbels_gebruiken=False):
    """Haal inschrijvingen op voor een lijst tornooien via de VTTL API.

    Vraagt per tornooi de registraties op, filtert op provincie en sluit
    dubbel/mixed/menu-reeksen standaard uit (tenzij ``dubbels_gebruiken=True``).
    De A-reeksen Dames en Heren zijn gratis en worden buiten de betalingslijst
    gehouden.  Optioneel worden handmatig ingegeven dubbels uit een Excel-bestand
    samengevoegd.

    Args:
        tornooien (list[str]): Lijst van tornooinamen zoals ze in de VTTL API
            voorkomen.
        inschrijvingsgeld (int | dict): Inschrijvingsgeld per tornooi.  Geef een
            enkel getal voor een vast bedrag, of een woordenboek
            ``{tornooinaam: bedrag}`` voor variabele bedragen.
        file_dubbels (str, optional): Pad naar een Excel-bestand met handmatig
            ingegeven dubbels.  Standaard ``'Dubbels.xlsx'``.
        provincie (str, optional): Provinciefilter op basis van het eerste teken
            van de clubcode.  Standaard ``'A'`` (Antwerpen).
        dubbels_gebruiken (bool, optional): Indien ``True`` worden ook
            dubbel/mixed-reeksen opgenomen.  Standaard ``False``.

    Returns:
        tuple[pandas.DataFrame, pandas.DataFrame]:
            - ``final``: Geaggregeerd DataFrame per speler met totale
              inschrijvingsgelden en een kommalijst van tornooien.
            - ``all_registrations``: Gedetailleerd DataFrame met één rij per
              inschrijving (voor rapportage en totalen).
    """
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
                # gratis in A-reeks dames en heren
                if not ((serie['Name']== 'Dames A')or(serie['Name']=='Heren A')):
                    # methode werkt momenteel niet voor dubbels
                
                    if (not any(sub in serie['Name'].lower() for sub in ['dubbel', 'doubles', 'mixed', 'mixte','Menu','menu']) or dubbels_gebruiken):
                        print('\t opgenomen reeks:',serie['Name'])    
                        nummers=[w['Member']['UniqueIndex'] for w in serie['RegistrationEntries']]
                        nummers=pd.DataFrame(nummers,columns=['Lidnummer'])
                        nummers['Tornooi']=item['Name'] if item['Name'] != 'BK A eindtabellen - CB A tableaux final' else 'BK A - CB A'
                        nummers['reeks']=item['Name']+' '+serie['Name']
                        nummers['Naam']=[w['Member']['LastName'] for w in serie['RegistrationEntries']]
                        nummers['Voornaam']=[w['Member']['FirstName'] for w in serie['RegistrationEntries']]
                        nummers['Club']=[w['Club']['UniqueIndex'] for w in serie['RegistrationEntries']]
                        

                        # inschrijvingsgeld per tornooi
                        if isinstance(inschrijvingsgeld, int):
                            nummers['Inschrijvingsgeld']=inschrijvingsgeld
                        else:
                            nummers['Inschrijvingsgeld']=inschrijvingsgeld[item['Name']]

                        dfs.append(nummers)
                    else:
                        print('\t Niet opgenomen reeks:',serie['Name'])


    #combine all
    all_registrations=pd.concat(dfs)

    # only one participation in the same tournament
    all_registrations=all_registrations.drop_duplicates(['Lidnummer','Naam','Voornaam','Club','Tornooi'])

    # select province
    all_registrations=all_registrations[[w[0]=='A' for w in all_registrations.Club]]
    all_registrations=all_registrations[all_registrations['Club']!='AFTT']

    # add dubbels
    # if dubbesl exists, add them
    if os.path.exists(file_dubbels):
        tmp=pd.read_excel(file_dubbels)
        all_registrations=pd.concat([all_registrations,tmp])

    # all Name and Voornaam to capital letter for first letter of each word
    all_registrations['Naam'] = all_registrations['Naam'].str.title()
    all_registrations['Voornaam'] = all_registrations['Voornaam'].str.title()


    # remove leading and trailing spaces
    all_registrations['Naam']=all_registrations['Naam'].str.strip()
    all_registrations['Voornaam']=all_registrations['Voornaam'].str.strip()

    # add price in name of tournament
    all_registrations['Tornooi']=all_registrations['Tornooi']+' (EUR '+all_registrations['Inschrijvingsgeld'].astype(str)+')'

    # convert to simple list
    combi_amount=all_registrations.groupby(['Lidnummer','Naam','Voornaam','Club'])[['Inschrijvingsgeld']].sum()
    combi_list_tournaments=all_registrations.groupby(['Lidnummer','Naam','Voornaam','Club'])['Tornooi'].apply(list).apply(lambda p:','.join(p))
    final=combi_amount.merge(combi_list_tournaments,left_index=True,right_index=True).reset_index()

    # check on duplicated lidnummer
    if final['Lidnummer'].duplicated().sum()>0:
        print('Duplicated lidnummer, stopping')
        print(final[final['Lidnummer'].duplicated()])
        sys.exit()
    
    return final,all_registrations

def AddFines(final, file_boetes, tag_lidnummer='Lidnummer', tag_supplement='Supplement'):
    """Voeg boetes (niet-deelname) toe aan de inschrijvingslijst.

    Laadt een Excel-boetelijst in en koppelt deze op lidnummer aan het
    inschrijvings-DataFrame.  Boetes voor spelers die niet ingeschreven zijn,
    worden als extra rijen toegevoegd.  Een kolom ``'Totaal'`` wordt berekend
    als de som van inschrijvingsgeld en supplement.

    Args:
        final (pandas.DataFrame): Inschrijvings-DataFrame zoals geretourneerd
            door ``GetTournamentEntries()``.
        file_boetes (str): Pad naar het Excel-bestand met boetes.  Verwachte
            kolommen: ``Lidnummer``, ``Supplement``, ``Reden Supplement``.
        tag_lidnummer (str, optional): Kolomnaam voor het lidnummer.
            Standaard ``'Lidnummer'``.
        tag_supplement (str, optional): Kolomnaam voor het supplement/boete.
            Standaard ``'Supplement'``.

    Returns:
        pandas.DataFrame: Het bijgewerkte DataFrame met kolommen
        ``Supplement``, ``Reden Supplement`` en ``Totaal``.
    """
    # load input
    boetes=pd.read_excel(file_boetes)

    # merge op lidnummer
    final=final.merge(boetes[[tag_lidnummer,tag_supplement,'Reden Supplement']],left_on=tag_lidnummer,right_on=tag_lidnummer,how='left')

    # check op boete-lijnen die niet gematcht zijn en voeg ze toe
    unmatched_boetes=boetes.merge(
        final[[tag_lidnummer]].drop_duplicates(),
        on=tag_lidnummer,
        how='left',
        indicator=True
    )
    unmatched_boetes=unmatched_boetes[unmatched_boetes['_merge']=='left_only']

    if len(unmatched_boetes)>0:
        print('\nBoetes zonder match in final (worden toegevoegd):')
        cols_to_print=[w for w in ['Naam','Voornaam',tag_lidnummer,tag_supplement,'Reden Supplement'] if w in unmatched_boetes.columns]
        print(unmatched_boetes[cols_to_print])

        add_rows=pd.DataFrame(columns=final.columns)

        # kopieer beschikbare velden vanuit boetes
        for col in unmatched_boetes.columns:
            if col in add_rows.columns and col!='_merge':
                add_rows[col]=unmatched_boetes[col].values

        # standaardwaarden voor niet-ingeschreven spelers met boete
        if 'Inschrijvingsgeld' in add_rows.columns:
            add_rows['Inschrijvingsgeld']=0
        if 'Tornooi' in add_rows.columns:
            add_rows['Tornooi']='Geen'
        if 'Club' in add_rows.columns:
            add_rows['Club']=add_rows['Club'].fillna('')

        final=pd.concat([final,add_rows],ignore_index=True)

    final['Totaal']=final['Inschrijvingsgeld']+final[tag_supplement].fillna(0)
    
    return final
    
def PrintTotals(final, all_registrations, uitvoer_detail_totaal, column_supplement='Supplement', column_reason_supplement='Reden Supplement'):
    """Druk totaaloverzichten af en sla het gedetailleerde saldo op als Excel.

    Toont via stdout de inschrijvingsgelden per tornooi en de supplementen per
    reden.  Slaat vervolgens het volledige saldo-DataFrame op als Excel-bestand.

    Args:
        final (pandas.DataFrame): Geaggregeerd inschrijvings-DataFrame (output
            van ``AddFines``).
        all_registrations (pandas.DataFrame): Gedetailleerde inschrijvingslijst
            (output van ``GetTournamentEntries``).
        uitvoer_detail_totaal (str): Pad naar het te schrijven Excel-bestand.
        column_supplement (str, optional): Kolomnaam voor het supplement.
            Standaard ``'Supplement'``.
        column_reason_supplement (str, optional): Kolomnaam voor de reden van
            het supplement.  Standaard ``'Reden Supplement'``.
    """
    
    print('\nTotalen')
    print(all_registrations.groupby('Tornooi').sum()['Inschrijvingsgeld'])
    
    print('\nTotale supplementen')
    print(final.groupby(column_reason_supplement).sum()[column_supplement])

    # opslaan lijst nog te betalen
    df=final.sort_values('Naam')[['Lidnummer', 'Naam','Voornaam','Club', 'Totaal','Inschrijvingsgeld', column_supplement,
          'Tornooi',column_reason_supplement]].rename(columns={'Totaal':'Totaal (euro)'})
    SaveExcel(df,uitvoer_detail_totaal,'Saldo')




def TournamentsSaveAndMail(final, clubs_direct, default_items, uitvoer_detail_club, uitvoer_detail_individueel, lege_factuur, tornooien, startnummer_factuur=0, formaat_factuur='A-2023/', column_supplement='Supplement', column_reason_supplement='Reden Supplement', functies=['secretaris', 'voorzitter', 'penningmeester'], send_mails=False, mail_test=True):
    """Sla tornooisaldi op als Excel en Word-facturen en verstuur e-mails.

    Verwerkt twee groepen clubs afzonderlijk:

    1. **Individuele spelers** (clubs die niet rechtstreeks betalen): per speler
       wordt een e-mail samengesteld met het persoonlijke saldo.  De gegevens
       worden opgeslagen in ``uitvoer_detail_individueel``.

    2. **Directe club-afrekening**: voor clubs in ``clubs_direct`` worden een
       gezamenlijk Excel-bestand (``uitvoer_detail_club``) en een Word-factuur
       per club aangemaakt.  Bij ``send_mails=True`` worden factuur + Excel als
       bijlage gemaild.

    Args:
        final (pandas.DataFrame): Geaggregeerd inschrijvings-DataFrame inclusief
            kolommen ``Totaal``, ``Inschrijvingsgeld``, ``Supplement``.
        clubs_direct (list[str]): Lijst van clubcodes die rechtstreeks per club
            betalen.
        default_items (dict): Woordenboek met tekstsjablonen en vervangwaarden
            voor de Word-factuur en e-mailberichten.  Verwachte sleutels:
            ``'Mailindividueel'``, ``'MailClubs'``, en vervangingssleutels als
            ``'DATUM'``, ``'REKENING'``, etc.
        uitvoer_detail_club (str): Pad naar het Excel-bestand voor de
            club-afrekening.
        uitvoer_detail_individueel (str): Pad naar het Excel-bestand voor de
            individuele afrekening.
        lege_factuur (str): Pad naar het lege Word-sjabloon voor de factuur.
        tornooien (list[str]): Lijst van tornooinamen (voor vermelding in de
            factuur).
        startnummer_factuur (int, optional): Startnummer voor factuurnummering.
            Standaard ``0``.
        formaat_factuur (str, optional): Prefix voor het factuurnummer.
            Standaard ``'A-2023/'``.
        column_supplement (str, optional): Kolomnaam voor supplementen.
            Standaard ``'Supplement'``.
        column_reason_supplement (str, optional): Kolomnaam voor de reden van
            het supplement.  Standaard ``'Reden Supplement'``.
        functies (list[str], optional): Clubfuncties voor mailadressen.
            Standaard ``['secretaris', 'voorzitter', 'penningmeester']``.
        send_mails (bool, optional): Indien ``True`` worden e-mails verstuurd.
            Standaard ``False``.
        mail_test (bool, optional): Indien ``True`` worden mails in testmodus
            verzonden (zie ``send_emails``).  Standaard ``True``.
    """
    
    
    # selectie clubs die niet rechtsteeks betalen
    #############################################
    print('\nAfrekening per speler')
    selectie_betalen=final[~final.Club.isin(clubs_direct)]
    if len(selectie_betalen)>0:
        # mailadressen
        mailadressen=mailroutines.GetMailinglijst_Lidnummer(selectie_betalen[['Lidnummer']],functies=functies)

        # opslaan lijst in Excel
        df=selectie_betalen.sort_values('Naam')[['Lidnummer', 'Naam','Voornaam','Club', 'Totaal','Inschrijvingsgeld', column_supplement,
            'Tornooi',column_reason_supplement]].rename(columns={'Totaal':'Totaal (euro)'})
        
        
        start_mail_send=df.merge(mailadressen)
        SaveExcel(start_mail_send,uitvoer_detail_individueel,'Saldo')

        # vervang een aantal trefwoorden in mail_in
        mail_in=default_items['Mailindividueel']
        start_mail_send.rename(columns={'Email':'receiver'},inplace=True)
        start_mail_send['subject']='Inschrijvingsgeld tornooien tafeltennis '+start_mail_send['Voornaam']+' '+start_mail_send['Naam']
        start_mail_send[column_reason_supplement] = start_mail_send[column_reason_supplement].apply(lambda x: f'({x})' if pd.notnull(x) else '')
        start_mail_send[column_supplement] = start_mail_send[column_supplement].fillna(0)
        start_mail_send['message']=start_mail_send.apply(lambda row: mail_in.replace('TOTAAL',str(row['Totaal (euro)'])).replace('INSCHRIJVINGSGELDEN',str(row['Inschrijvingsgeld'])).replace('TORNOOIEN',row['Tornooi']).replace('BOETES',str(row[column_supplement])).replace('REDENBOETE',row[column_reason_supplement]).replace('SPELER',row['Voornaam']+' '+row[ 'Naam']),axis=1)

        # stuur mails uit
        if send_mails:

            mailroutines.send_emails(start_mail_send,smtp_server='mail.tafeltennisantwerpen.be',smtp_port=587,test_mode=mail_test)

    # verwerking clubs die rechtsteeks betalen
    #############################################
    print('\nAfrekening per club, direct betaling')
    selectie_betalen_club=final[final.Club.isin(clubs_direct)]

    # mailadressen
    mailclubs,club_details=mailroutines.GetMailClubs(selectie_betalen_club.Club.unique(),functies=functies)
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
    docs=dict()
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

            docs[clubnummer]=os.path.dirname(uitvoer_detail_club)+'/'+clubnummer+'_'+str(startnummer_factuur+cnt)+'.docx'
            doc.save(docs[clubnummer])
            
            cnt=cnt+1

            # save table            
            SaveExcel(sel[['Naam','Voornaam','Totaal','Inschrijvingsgeld',column_supplement,column_reason_supplement,'Tornooi']],docs[clubnummer].replace('docx','xlsx'),'Saldo')

    
    # send emails
    if send_mails:
        club_details['subject']='Inschrijvingsgeld tornooien tafeltennis '+club_details.index
        club_details['message']=default_items['MailClubs']
        club_details['receiver']=club_details['Email']
        club_details['attachment']=[docs[w]+','+docs[w].replace('docx','xlsx') for w in club_details.index]
        mailroutines.send_emails(club_details,smtp_server='mail.tafeltennisantwerpen.be',smtp_port=587,test_mode=mail_test)


def Inschrijving(lidnummer, tornooi, reeks=None, mail=True, dubbel=False, unregister=False, check=False):
    """Schrijf een speler in voor een tornooi via de VTTL API (op basis van lidnummer).

    Args:
        lidnummer (int | list[int]): Lidnummer van de speler.  Bij een dubbel
            een lijst van twee lidnummers.
        tornooi (str): Naam van het tornooi zoals vermeld in de VTTL API.
        reeks (str, optional): Naam van de reeks.  Verplicht; het script stopt
            als dit niet opgegeven is.
        mail (bool, optional): Indien ``True`` ontvangt de speler een
            bevestigingsmail van de VTTL API.  Standaard ``True``.
        dubbel (bool, optional): Indien ``True`` wordt een dubbelinschrijving
            gedaan; ``lidnummer`` moet dan een lijst zijn.  Standaard ``False``.
        unregister (bool, optional): Indien ``True`` wordt de inschrijving
            ongedaan gemaakt.  Standaard ``False``.
        check (bool, optional): Indien ``True`` wordt de API-aanroep
            overgeslagen (droge run).  Standaard ``False``.

    Raises:
        ValueError: Als ``dubbel=True`` maar ``lidnummer`` geen lijst is, of
            omgekeerd.
        zeep.exceptions.Fault: Bij een API-fout van de VTTL-server.
    """

    export, clubs = mailroutines.getExport()
    if isinstance(lidnummer, int):
        persoon = str(export[export['Lidnummer'] == lidnummer]['Naam'].iloc[0])
    else:
        persoon = ','.join(list(export[export['Lidnummer'].isin(lidnummer)]['Naam']))

    client, input_dict = _create_client()
    account, paswoord = _get_credentials()
    index_tornooi, index_reeks = _lookup_tournament_indices(input_dict, tornooi, reeks)

    try:
        if dubbel:
            if isinstance(lidnummer, list):
                if not check:
                    response = client.service.TournamentRegister(TournamentUniqueIndex=index_tornooi, SerieUniqueIndex=index_reeks, PlayerUniqueIndex=lidnummer, Credentials={'Account': account, 'Password': paswoord}, NotifyPlayer=mail, Unregister=unregister)
            else:
                raise ValueError("Inschrijving dubbel: lidnummer is geen lijst")
        else:
            if isinstance(lidnummer, int):
                if not check:
                    response = client.service.TournamentRegister(TournamentUniqueIndex=index_tornooi, SerieUniqueIndex=index_reeks, PlayerUniqueIndex=str(lidnummer), Credentials={'Account': account, 'Password': paswoord}, NotifyPlayer=mail, Unregister=unregister)
            else:
                raise ValueError("Inschrijving enkel: lijsten van lidnummers gegeven")

        print('Inschrijving voor ', persoon, ' (', lidnummer, ') in ', tornooi, ' in reeks ', reeks, ' is gelukt.')
    except ValueError as ve:
        print(f"Caught an error: {ve}")
    except zeep.exceptions.Fault as fault:
        print(f"Caught a zeep Fault: {fault}")


def InschrijvingNaam(naam, tornooi, reeks=None, mail=True, dubbel=False, unregister=False, check=False):
    """Schrijf een speler in voor een tornooi via de VTTL API (op basis van naam).

    Zoekt het lidnummer op via de VTTL-ledenexport en delegeert vervolgens
    naar de SOAP-API.

    Args:
        naam (str | list[str]): Volledige naam (achternaam voornaam) van de
            speler.  Bij een dubbel een lijst van twee namen.
        tornooi (str): Naam van het tornooi zoals vermeld in de VTTL API.
        reeks (str, optional): Naam van de reeks.  Verplicht.
        mail (bool, optional): Indien ``True`` ontvangt de speler een
            bevestigingsmail.  Standaard ``True``.
        dubbel (bool, optional): Indien ``True`` wordt een dubbelinschrijving
            gedaan; ``naam`` moet dan een lijst zijn.  Standaard ``False``.
        unregister (bool, optional): Indien ``True`` wordt de inschrijving
            ongedaan gemaakt.  Standaard ``False``.
        check (bool, optional): Indien ``True`` wordt de API-aanroep
            overgeslagen (droge run).  Standaard ``False``.

    Raises:
        ValueError: Als ``dubbel=True`` maar ``naam`` geen lijst is, of
            omgekeerd.
        zeep.exceptions.Fault: Bij een API-fout van de VTTL-server.
    """

    export, clubs = mailroutines.getExport()
    if isinstance(naam, list):
        lidnummer = [export[export['naam_combi'] == w.lower()]['Lidnummer'].iloc[0] for w in naam]
    else:
        lidnummer = export[export['naam_combi'] == naam.lower()]['Lidnummer'].iloc[0]

    client, input_dict = _create_client()
    account, paswoord = _get_credentials()
    index_tornooi, index_reeks = _lookup_tournament_indices(input_dict, tornooi, reeks)

    try:
        if dubbel:
            if isinstance(naam, list):
                if not check:
                    response = client.service.TournamentRegister(TournamentUniqueIndex=index_tornooi, SerieUniqueIndex=index_reeks, PlayerUniqueIndex=lidnummer, Credentials={'Account': account, 'Password': paswoord}, NotifyPlayer=mail, Unregister=unregister)
            else:
                raise ValueError("Inschrijving dubbel: naam is geen lijst")
        else:
            if isinstance(naam, str):
                if not check:
                    response = client.service.TournamentRegister(TournamentUniqueIndex=index_tornooi, SerieUniqueIndex=index_reeks, PlayerUniqueIndex=str(lidnummer), Credentials={'Account': account, 'Password': paswoord}, NotifyPlayer=mail, Unregister=unregister)
            else:
                raise ValueError("Inschrijving enkel: lijsten van namen gegeven")

        print('Inschrijving voor ', naam, ' (', lidnummer, ') in ', tornooi, ' in reeks ', reeks, ' is gelukt.')
    except ValueError as ve:
        print(f"Caught an error: {ve}")
    except zeep.exceptions.Fault as fault:
        print(f"Caught a zeep Fault: {fault}")
