import pandas as pd


def loadData():
    from ttpy.mailroutines import getExport

    leden=getExport()[0]

    return leden
    
    
def checkall():
    checkStatuten()
    checkGeboortedatum()
    checkDamesOpHeren()
    




def checkStatuten():

    print('\n\n')
    print('--------------')
    print('Check statuten')
    print('--------------')
    
    # load data
    data=loadData()
    
    # leden zonder statuut
    selectie=data[data['Statuut (0)'].isnull()]
    
    print('=> Leden zonder statuut')
    for club in selectie['Club (0)'].unique():
        sel2=selectie[selectie['Club (0)']==club]
        print('\n')
        print(club,sel2['NaamClub'].unique())
        print(sel2[['Naam','Voornaam','Actief Start Datum (0)','Gebooredatum']])


    print('\n')
    print('=> Statuut jeugdspelers')
    print(data[(data.Leeftijdscategorie.isin(['Kadet','Miniem','Benjamin','Preminiem']))&(data['Statuut (0)']!='competitief lid')][['Naam','Voornaam','Club (0)','NaamClub','Leeftijdscategorie','Statuut (0)']])


def checkGeboortedatum():
    print('\n\n')
    print('--------------')
    print('Check leeftijd')
    print('--------------')
    
    # load data
    data=loadData()
    
    # jaartal
    from datetime import datetime
    current_year = datetime.now().year
    start_check=pd.to_datetime(str(current_year-3)+'-01-01',dayfirst=False)

    print('\n')
    print('Geboortedatum')
    print(data[pd.to_datetime(data.Gebooredatum)>start_check][['Naam','Voornaam','Club (0)','NaamClub','Gebooredatum','Actief Start Datum (0)']])

def checkDamesOpHeren():

    print('\n\n')
    print('--------------------')
    print('Check dames op heren')
    print('--------------------')
    

    # load data
    data=loadData()
    
    # checks
    selectie=data[(data.Geslacht=='V')&(data['Statuut (0)'].isin(['competitief lid', 'recreant-reserve']))&(data['Geen Matchen Heren']==0)]
    controle=selectie.groupby(['Club (0)','NaamClub'])[['Naam']].count()
    controle.rename(columns={'Naam':'AantalDamesOpHerenSterktelijst'},inplace=True)
    probleem=controle[controle['AantalDamesOpHerenSterktelijst']>8]
    print('Clubs met meer dan 8 dames:\n',probleem)
    
    print('\n\nOverzicht per club:')
    probleem=probleem.reset_index()

    for item in probleem['Club (0)']:
        print(item)
        print(selectie[selectie['Club (0)']==item][['Naam','Voornaam','Geen Matchen Heren','Statuut (0)','Leeftijdscategorie','Klassement Heren (0)','Club Start Datum (0)']])
        print(selectie[selectie['Club (0)']==item][['Naam','Voornaam']].to_string(index=False))
        print('\n')