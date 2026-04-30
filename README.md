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
Elke tool leest zijn invoer uit een **TOML-configuratiebestand**.  
Voorbeeldconfigs staan in de `configs/` map van het package.  
Standaard wordt gezocht naar `<naam>.toml` in de huidige map; een alternatief pad opgeven kan via `--config`.

```bash
ttpy_mailing_naam --config mijn_config.toml
```

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

Voert alle ledencontroles uit en print de resultaten naar de console.

- **Statuten** — leden zonder statuut, jeugdspelers met verkeerd statuut, recreanten met te hoog klassement
- **Geboortedatum** — leden met een geboortedatum die te recent lijkt
- **Dames op herensterktelijst** — clubs met meer dan 8 dames op de herenlijst

**Geen configuratiebestand nodig.**

```bash
ttpy_check
```

---

## Module-overzicht

| Module | Inhoud |
|---|---|
| `mailroutines` | Ledenexport inlezen, e-mailadressen opzoeken, mails versturen |
| `tournamentroutines` | Tornooi-inschrijvingen ophalen, verwerken en factureren |
| `generalroutines` | Hulpfuncties voor Word- en Excel-bestanden |
| `check_routines` | Kwaliteitscontroles op ledendata |
| `vttl_download` | Automatisch downloaden van de VTTL ledenexport |

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
