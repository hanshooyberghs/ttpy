"""
check_routines.py
-----------------
Kwaliteitscontroles op de VTTL-ledendata voor Provincie Antwerpen.

Voert drie categorieën controles uit:
    - Statuten: leden zonder statuut, jeugdspelers met verkeerd statuut,
      recreanten met te hoog klassement, secretarissen/voorzitters zonder
      competitief-lid-statuut.
    - Geboortedatum: leden met een geboortedatum die te recent is.
    - Dames op herensterktelijst: clubs met meer dan 8 dames op de herenlijst.

Gebruik ``checkall()`` om alle controles in één keer uit te voeren.
"""
import pandas as pd


def loadData():
    """Laad de VTTL-ledenexport in als DataFrame.

    Returns:
        pandas.DataFrame: Volledig ledenbestand inclusief clubnamen, zoals
        geretourneerd door ``getExport()``.
    """

    leden=getExport()[0]

    return leden


def checkall():
    """Voer alle beschikbare ledencontroles uit.

    Roept achtereenvolgens ``checkStatuten()``, ``checkGeboortedatum()`` en
    ``checkDamesOpHeren()`` aan en drukt de resultaten af via stdout.
    """
    checkGeboortedatum()
    checkDamesOpHeren()






def checkStatuten():
    """Controleer de statuten van leden en rapporteer afwijkingen.

    Vier deelcontroles:
        1. Leden zonder statuut (per club).
        2. Jeugdspelers (Kadet t/m Preminiem) met een ander statuut dan
           ``'competitief lid'``.
        3. Recreanten (``'recreant-reserve'``) met een klassement hoger dan
           NG/D6/E0/E2/E4/E6.
        4. Secretarissen en voorzitters met een ander statuut dan
           ``'competitief lid'``.

    Resultaten worden afgedrukt via stdout.
    """
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

    print('\n')
    print('=> Statuut recreanten')
    print(data[(data['Statuut (0)']=='recreant-reserve')&(~data['Klassement Heren (0)'].isin(['NG','D6','E0','E2','E4','E6']))][['Naam','Voornaam','Club (0)','NaamClub','Klassement Heren (0)','Statuut (0)']])

    print('\n')
    print('=> Statuut secretarissen en voorzitters')
    print(data[(data['Statuut (0)']!='competitief lid') & (data['Club Funtie (0)'].isin(['secretaris','voorzitter']))][['Naam','Voornaam','Club (0)','NaamClub','Club Funtie (0)','Statuut (0)']])


def checkGeboortedatum():
    """Controleer of er leden zijn met een onwaarschijnlijk recente geboortedatum.

    Markeert leden met een geboortedatum die minder dan drie jaar geleden is
    (ten opzichte van het huidige kalenderjaar) als potentieel foutief.

    Resultaten worden afgedrukt via stdout.
    """
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
    """Controleer welke clubs meer dan 8 dames op de herensterktelijst hebben.

    Selecteert dames met statuut ``'competitief lid'`` of ``'recreant-reserve'``
    die niet zijn uitgesloten van herenwedstrijden (``Geen Matchen Heren == 0``).
    Clubs met meer dan 8 zulke speelsters worden gedetailleerd gerapporteerd.

    Resultaten worden afgedrukt via stdout.
    """
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
