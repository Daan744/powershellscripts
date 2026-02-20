# PowerShell Scripts

Deze repository bevat twee beheerscripts voor Microsoft 365 en Active Directory.

## Inhoud

- **`Create365User.ps1`**  
  GUI-script voor het aanmaken van nieuwe gebruikers in:
  - **On-Premises AD + Microsoft 365** (hybride scenario)
  - **Cloud Only (Microsoft 365 / Entra ID)**

- **`Copy-ADUserGroupMembership.ps1`**  
  CLI-script om groepslidmaatschappen van één AD-gebruiker naar een andere AD-gebruiker te kopiëren.

---

## Vereisten

### Algemeen

- Windows PowerShell 5.1 (aanbevolen voor WinForms en AD-module)
- Voldoende rechten in Active Directory en/of Microsoft 365

### Modules

Afhankelijk van het gekozen script/scenario:

- `ActiveDirectory`
- `Microsoft.Graph` (voor Graph login en M365 acties)
- `ExchangeOnlineManagement` (waar nodig voor mailbox/licentieflow)

> `Create365User.ps1` probeert ontbrekende modules tijdens runtime te laden/installeren.

---

## Script 1: `Create365User.ps1`

### Wat doet dit script?

Dit script start een Windows GUI waarmee je stapsgewijs een gebruiker aanmaakt. Het ondersteunt zowel on-prem als cloud-only provisioning en haalt tenantinformatie (zoals domeinen/licenties) dynamisch op.

### Starten

Open PowerShell en voer uit vanuit de repo-map:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Create365User.ps1
```

### Opmerkingen

- Het script zet de `ExecutionPolicy` voor `CurrentUser` naar `RemoteSigned` als dat nodig is.
- Er wordt een logbestand aangemaakt in dezelfde map als het script (`Create365User.log`).
- Voor Graph-login wordt interactieve authenticatie gebruikt.

---

## Script 2: `Copy-ADUserGroupMembership.ps1`

### Wat doet dit script?

Kopieert AD-groepslidmaatschappen van een brongebruiker naar een doelgebruiker, zonder dubbele toevoegingen.

### Voorbeeldgebruik

**Dry-run met `-WhatIf` (aanbevolen):**

```powershell
.\Copy-ADUserGroupMembership.ps1 -SourceUser jansenj -TargetUser peetersm -WhatIf
```

**Echt uitvoeren:**

```powershell
.\Copy-ADUserGroupMembership.ps1 -SourceUser jansenj -TargetUser peetersm
```

**Zonder standaardgroep `Domain Users`:**

```powershell
.\Copy-ADUserGroupMembership.ps1 -SourceUser jansenj -TargetUser peetersm -SkipDefaultPrimaryGroup
```

### Parameters

- `-SourceUser` (verplicht): brongebruiker (SAM/account-identiteit)
- `-TargetUser` (verplicht): doelgebruiker
- `-SkipDefaultPrimaryGroup` (optioneel): slaat `Domain Users` over
- `-WhatIf` (standaard PowerShell `ShouldProcess`): toont wat er zou gebeuren

---

## Veilig gebruik / best practices

1. Test altijd eerst met `-WhatIf` (voor groepswijzigingen).
2. Gebruik een account met minimale noodzakelijke rechten.
3. Controleer logs en output na uitvoering.
4. Test nieuwe tenant- of OU-scenario’s eerst in een testomgeving.

---

## Troubleshooting

- **Module niet gevonden**: installeer/importeer de vereiste PowerShell-modules.
- **Toegang geweigerd**: controleer RBAC/AD-permissies.
- **Graph-login faalt**: controleer Conditional Access/MFA en admin-consent voor benodigde Graph-permissies.

---

## Disclaimer

Gebruik deze scripts op eigen risico. Test altijd in een niet-productieomgeving voordat je ze in productie inzet.
