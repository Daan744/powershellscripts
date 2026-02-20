# PowerShell Scripts

Deze repository bevat twee PowerShell-scripts voor Microsoft 365 en Active Directory.

## Scripts (direct kopieerbare raw links)

| Script | Doel | Raw link |
|---|---|---|
| `Create365User.ps1` | GUI voor het aanmaken van gebruikers (On-Prem + M365 of Cloud Only). | https://raw.githubusercontent.com/Daan744/powershellscripts/refs/heads/main/Create365User.ps1 |
| `Copy-ADUserGroupMembership.ps1` | Kopieert AD-groepslidmaatschappen van brongebruiker naar doelgebruiker. | https://raw.githubusercontent.com/Daan744/powershellscripts/refs/heads/main/Copy-ADUserGroupMembership.ps1 |

---

## Vereisten

- Windows PowerShell 5.1 (aanbevolen)
- Voldoende rechten in Active Directory en/of Microsoft 365
- Modules (afhankelijk van scenario):
  - `ActiveDirectory`
  - `Microsoft.Graph`
  - `ExchangeOnlineManagement`

---

## Gebruik

### 1) Create365User.ps1

Start vanuit de repo-map:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Create365User.ps1
```

### 2) Copy-ADUserGroupMembership.ps1

Dry-run (aanbevolen):

```powershell
.\Copy-ADUserGroupMembership.ps1 -SourceUser jansenj -TargetUser peetersm -WhatIf
```

Echt uitvoeren:

```powershell
.\Copy-ADUserGroupMembership.ps1 -SourceUser jansenj -TargetUser peetersm
```

Optioneel zonder `Domain Users`:

```powershell
.\Copy-ADUserGroupMembership.ps1 -SourceUser jansenj -TargetUser peetersm -SkipDefaultPrimaryGroup
```

---

## Veilig gebruik

1. Gebruik eerst `-WhatIf` bij groepswijzigingen.
2. Werk met minimale rechten.
3. Test wijzigingen eerst in een testomgeving.

---

## Disclaimer

Gebruik deze scripts op eigen risico.
