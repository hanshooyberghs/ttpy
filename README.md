# ttpy — Tafeltennis Provincie Antwerpen · Python package

Python-hulppakket voor het verwerken van VTTL-ledendata en het beheer van provinciale tornooien.

---

## Installatie

```bash
# Eenmalig, in de ttpy conda-omgeving
conda activate ttpy
pip install -e /pad/naar/ttpy
```

Vereiste omgevingsvariabelen (stel in via `.env`, shell-profiel of het besturingssysteem):

| Variabele | Inhoud |
|---|---|
| `EXPORT_TT` | Pad naar de VTTL ledenexport CSV (`exportPerson.csv`) |
| `PASWOORD_TT` | Wachtwoord voor leden.vttl.be en VTTL API |
| `ACCOUNT_TT` | Gebruikersnaam voor de VTTL API |
| `MAIL_PASSWORD` | Wachtwoord voor het SMTP/IMAP mailaccount |

---

## Command-line tools

Alle commando's zijn beschikbaar na `conda activate ttpy`.  
Configureerbare tools lezen hun invoer uit een **TOML-configuratiebestand**.
Voorbeeldconfigs staan in de `configs/` map van het package.  
Bij de bestaande configureerbare tools wordt standaard gezocht naar `<naam>.toml`
in de huidige map; een alternatief pad kan via `--config`. De ledenaanpassing
vereist altijd expliciet `--config`.

```bash
ttpy_mailing_naam --config mijn_config.toml
```

---

### `ttpy_ledenlijst_aanpassing`

Wijzigt uitsluitend persoonsvelden en statuut en kan expliciet een **eerste**
klassement toevoegen. Geen export, mail of andere relaties. `--config` is verplicht.
Een blanco, volledig becommentarieerde template staat in de werkmap
`/mnt/g/My Drive/TT Provincie/ttpy_configs/ledenlijst_aanpassing.toml`.
Kopieer deze voor een opdracht, verwijder `leden = []`, activeer `[[leden]]`
en alleen de gewenste velden. Herhaal het ledenblok voor meerdere leden.

```bash
ttpy_ledenlijst_aanpassing --config "/mnt/g/My Drive/TT Provincie/ttpy_configs/ledenlijst_aanpassing.toml" --dry-run
ttpy_ledenlijst_aanpassing --config mijn_aanpassingen.toml
```

Identificeer met `naam = "Achternaam Voornaam"` en optioneel `lidnummer`.
Voornaam/achternaam wijzigen gebeurt apart onder `[leden.wijzigingen]`.
Er moet exact één online match zijn; vergelijking negeert hoofdletters,
accenten en extra witruimte. Zonder lidnummer wordt het volledige, gepagineerde
ledenoverzicht (ook inactieve leden) gelezen: dit kan aanzienlijk langer duren.
Met lidnummer blijft de naamcontrole verplicht.

Alleen ingevulde waarden worden verwerkt: ontbrekende velden, `""` en `[]`
blijven ongemoeid; `false` en numerieke `0` worden wél verwerkt. Numerieke
adres/telefoon/identificatievelden mogen tekst of integers zijn; gebruik tekst
voor voorloopnullen. Een ingevulde categorieënlijst vervangt de bestaande selectie.
Dropdownwaarde `geen` is geldig. Onbekende keys, dubbele records/leden/velden,
onjuiste types, maximumlengtes, enums en datums worden gecontroleerd.
Datums zijn strings `DD/MM/YYYY`.

`statuut` staat direct onder het ledenblok. Klassement staat onder
`[leden.klassement]` met `heren`, `dames` en optioneel `begindatum` (standaard
vandaag, expliciet zichtbaar in het voorstel). Een damesklassement vereist
geslacht `vrouw`. Zodra **enig** online klassementrecord bestaat, inclusief
historische records, blokkeert een klassementaanvraag de hele uitvoering
vóór enige opslag. Er is geen vrije leegmaakfunctie.

Alle items worden vooraf online gecontroleerd. Live toont de tool het volledige
voorstel en vraagt eenmaal letterlijk `JA`; er is geen `--yes`.
`--dry-run` vraagt geen bevestiging en verstuurt geen opslagverzoeken.
Voor iedere opslag worden de precondities herlezen. Na iedere opslag worden
bedoelde en ongemoeide persoonsvelden, statuut, klassement en relaties online
geverifieerd. Bij een fout stopt de tool met een overzicht van geverifieerde
opslagen en mogelijk onzekere resultaten; **geen rollback of automatische retry**.
Ook een serverantwoord zonder HTTP-fout is pas na herlezing een succes.

Authenticatie gebruikt uitsluitend het bestaande VTTL-loginmechanisme en
`PASWOORD_TT` (of een veilige wachtwoordprompt); er is geen mailprompt.
Zet geen credentials in TOML. De blanco template (`leden = []`) doet niets,
zonder login. Een ledenblok met uitsluitend lege wijzigingswaarden verandert niets.

---

### `ttpy_trigger_download`

Logt in op [leden.vttl.be](https://leden.vttl.be), triggert een ledenexport, wacht op de e-mail met CSV-bijlagen en slaat die op. Vergelijkt automatisch met de vorige export en toont nieuwe en verwijderde leden.

**Geen configuratiebestand nodig.** Leest `EXPORT_TT`, `PASWOORD_TT` en `MAIL_PASSWORD` uit de omgeving.

```bash
ttpy_trigger_download
```

---

### `ttpy_mailing_naam`

Zoekt de e-mailadressen op van een lijst personen (op naam). Optioneel worden ook de clubcontacten (secretaris, voorzitter, …) van hun clubs toegevoegd.

**Configuratiebestand:** `mailing_naam.toml`

```toml
# functies: kies uit secretaris, voorzitter, penningmeester,
#           interclubSeniors, interclubJeugd, aanspreekpunt
# Lege lijst = enkel persoonlijke mailadressen
functies = []

# Namen: onder elkaar, komma-gescheiden of als TOML-array
namen = """
Peeters Jan
Maes Koen
De Smedt Lien
"""
```

```bash
ttpy_mailing_naam
ttpy_mailing_naam --config andere_lijst.toml
```

---

### `ttpy_mailing_clubs`

Zoekt e-mailadressen op van clubs via (een deel van) de clubnaam.

**Configuratiebestand:** `mailing_clubs.toml`

```toml
functies = "secretaris, voorzitter"

# Clubs: onder elkaar, komma-gescheiden of als TOML-array
clubs = "Turnhout, Rupel, Hoboken, Mechelen"
```

```bash
ttpy_mailing_clubs
```

---

### `ttpy_inschrijving_tornooi`

Schrijft spelers in (of uit) voor een tornooi via de VTTL API. Ondersteunt zowel enkelen als dubbels.

**Configuratiebestand:** `inschrijving_tornooi.toml`

```toml
tornooi    = "Provinciaal Kampioenschap Jeugd Antwerpen"
mail       = true      # stuur bevestigingsmail naar speler
unregister = false     # true = uitschrijven

# Enkeling
[[inschrijvingen]]
reeks  = "Heren B"
namen  = "Peeters Jan"
dubbel = false

# Dubbel: komma-gescheiden of lijst van 2 namen
[[inschrijvingen]]
reeks  = "Dubbel Gemengd -19"
namen  = "Michielsen Jente, Belmans Maithe"
dubbel = true
```

```bash
ttpy_inschrijving_tornooi
```

---

### `ttpy_mails_tornooi`

Haalt de mailinglijst op van clubs waarvan spelers ingeschreven zijn in één of meerdere tornooien.

**Configuratiebestand:** `mails_tornooi.toml`

```toml
functies = "secretaris"

# Tornooien: onder elkaar, komma-gescheiden of als TOML-array
tornooien = """
Provinciaal Kampioenschap Veteranen Antwerpen
Provinciaal Criterium Veteranen Antwerpen - Fase 1
"""
```

```bash
ttpy_mails_tornooi
```

---

### `ttpy_check`

Voert alle ledencontroles uit en print de resultaten naar de console. Daarna
worden de e-mailadressen per clubfunctie en als gecombineerde nieuwsbrieflijst
afgedrukt.

- **Statuten** — leden zonder statuut, jeugdspelers met verkeerd statuut, recreanten met te hoog klassement
- **Geboortedatum** — leden met een geboortedatum die te recent lijkt
- **Dames op herensterktelijst** — clubs met meer dan 8 dames op de herenlijst
- **Clubfuncties** — clubs met meer dan één voorzitter, secretaris of
  penningmeester; alle functievelden worden gecontroleerd en een lid telt
  per functie slechts eenmaal. Ontbrekende functies worden niet afgekeurd.
- **Mailinglijsten** — voorzitters, secretarissen, penningmeesters,
  interclubleiders, jeugdverantwoordelijken en alle functies gecombineerd

**Geen configuratiebestand nodig voor de bestaande controles.**

```bash
ttpy_check
```

Met een optionele `check.toml` wordt na download, controles en mailinglijsten
ook het jaarboek vernieuwd wanneer gepubliceerde bestuurs- of clubgegevens
veranderen. Vergelijking gebeurt met de JSON van de laatste succesvolle
publicatie, niet met CSV-bestandshashes: wijzigingen aan gewone leden leiden
niet tot een nieuwe PDF. Namen, functies, gepubliceerde contactgegevens,
clubnamen, lokalen en actieve clubs tellen wel mee. Bij een eerste publicatie,
een ontbrekende PDF of een gewijzigd seizoen wordt ook gegenereerd.
Een mislukte publicatie verandert de vergelijkingsbasis niet, zodat een
volgende uitvoering opnieuw probeert. Een ongeldig bronbestand geeft een fout.

```bash
ttpy_check --config "/pad/naar/TT Provincie/ttpy_configs/check.toml"
```

Configselectie: `--config`, daarna omgevingsvariabele `TTPY_CHECK_CONFIG`,
daarna `check.toml` in de huidige map. Een expliciet opgegeven maar ontbrekend
bestand geeft een fout; zonder config blijft de oude werking behouden.
Gebruik voor uitvoering vanuit wisselende mappen bij voorkeur
`TTPY_CHECK_CONFIG` met een absoluut pad in de ttpy-omgeving.

```toml
[jaarboek]
vaste_docx = "../Administratie/Jaarboek/Jaarboek_PCA_zonder_clubbesturen.docx"
club_pdf = "../Administratie/Jaarboek/Clubbesturen_Antwerpen.pdf"
jaarboek_pdf = "../Administratie/DocumentenPubliek/Jaarboek_PCA.pdf"
# seizoen = "2026-2027"  # anders uit de vaste DOCX
# vaste_pdf = "../Administratie/Jaarboek/Jaarboek_PCA_zonder_clubbesturen.pdf"
```

Alle relatieve paden zijn **relatief aan het configuratiebestand**.
Houd config en documenten buiten de package-installatie; een kopieerbare
template staat in `configs/check.toml`. Credentials blijven uitsluitend in de
omgeving. `EXPORT_TT` blijft de exportbron; een optionele `export`-override
moet bij `ttpy_check` naar dezelfde persoons-export verwijzen.
Zonder `vaste_pdf` wordt de DOCX met LibreOffice omgezet (Linux, Windows of
Windows vanuit WSL). Met `vaste_pdf` moet je die PDF zelf actueel houden.
De bewerkbare DOCX wordt nooit automatisch opnieuw aangemaakt of overschreven.
Paginanummers uit de vaste PDF worden niet herberekend.

### `ttpy_jaarboek`

Forceert de volledige publicatie met dezelfde config, zonder download of
ledencontroles. Gebruik dit na handmatige wijzigingen aan de vaste DOCX:
`ttpy_check` vergelijkt de gepubliceerde clubgegevens, niet DOCX-wijzigingen.

```bash
ttpy_jaarboek --config "/pad/naar/TT Provincie/ttpy_configs/check.toml"
```

De oorspronkelijke `Administratie/Jaarboek/maak_jaarboek.py` blijft als
compatibele ingang bestaan, met onder meer `--volledig` en `--maak-vaste-docx`.
De PDF-generator gebruikt `pypdf` en `lxml`; alleen de DOCX-conversie vereist
LibreOffice.

---

## Module-overzicht

| Module | Inhoud |
|---|---|
| `mailroutines` | Ledenexport inlezen, e-mailadressen opzoeken, mails versturen |
| `tournamentroutines` | Tornooi-inschrijvingen ophalen, verwerken en factureren |
| `generalroutines` | Hulpfuncties voor Word- en Excel-bestanden |
| `check_routines` | Kwaliteitscontroles op ledendata |
| `vttl_download` | Automatisch downloaden van de VTTL ledenexport |
| `jaarboek` | Club-PDF, DOCX-conversie, samenvoegen en wijzigingsdetectie |

---

## Configuratiebestanden — formaten

Lijstvelden in TOML-configs ondersteunen drie notaties:

```toml
# 1. TOML-array (vereist aanhalingstekens)
namen = ["Peeters Jan", "Maes Koen"]

# 2. Multiline string — geen aanhalingstekens nodig
namen = """
Peeters Jan
Maes Koen
"""

# 3. Komma-gescheiden string
namen = "Peeters Jan, Maes Koen"
```
