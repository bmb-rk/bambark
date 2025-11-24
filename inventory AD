<#
===========================================================
      AD INVENTORY COMPLETE – AUDIT D’ACCÈS & SERVEURS
===========================================================
#>

# --- 1) Charger le module AD ------------------------------------------
Import-Module ActiveDirectory

# --- 2) Fichier Excel de sortie ----------------------------------------
$Date = (Get-Date).ToString("yyyy-MM-dd")
$ExcelFile = "C:\AD_Inventory_$Date.xlsx"

# --- 3) Extraction : Utilisateurs AD -----------------------------------
$Users = Get-ADUser -Filter * -Properties * | Select `
Name, SamAccountName, UserPrincipalName, Enabled, Department, Title,
WhenCreated, PasswordLastSet, PasswordNeverExpires, LastLogonDate,
AccountExpirationDate, StreetAddress, City, Company, MemberOf

$Users | Export-Excel -Path $ExcelFile -WorksheetName "Users" -AutoSize

# --- 4) Extraction : Groupes AD ---------------------------------------
$Groups = Get-ADGroup -Filter * -Properties * | Select `
Name, GroupCategory, GroupScope, Description, WhenCreated

$Groups | Export-Excel -Path $ExcelFile -WorksheetName "Groups" -AutoSize

# --- 5) Membres des groupes sensibles ---------------------------------
$SensitiveGroups = @(
    "Domain Admins",
    "Enterprise Admins",
    "Schema Admins",
    "Administrators",
    "Backup Operators",
    "Account Operators",
    "Server Operators",
    "Remote Desktop Users"
)

foreach ($group in $SensitiveGroups) {
    if (Get-ADGroup -Filter "Name -eq '$group'") {
        $data = Get-ADGroupMember $group -Recursive | Select Name, SamAccountName, ObjectClass
        $ws = $group.Replace(" ", "_")
        $data | Export-Excel -Path $ExcelFile -WorksheetName $ws -AutoSize
    }
}

# --- 6) Extraction : Serveurs Windows ---------------------------------
$Servers = Get-ADComputer -Filter 'OperatingSystem -like "*Server*"'
        -Properties OperatingSystem, IPv4Address, LastLogonDate |
        Select Name, OperatingSystem, IPv4Address, LastLogonDate

$Servers | Export-Excel -Path $ExcelFile -WorksheetName "Servers" -AutoSize

# --- 7) Extraction : Accès RDP sur chaque serveur ----------------------
$RDPAccess = foreach ($server in $Servers.Name) {
    try {
        Get-WmiObject -Class Win32_GroupUser -ComputerName $server -ErrorAction Stop |
        Where-Object { $_.GroupComponent -like '*Remote Desktop Users*' } |
        Select @{N="Server";E={$server}}, *
    }
    catch {
        [PSCustomObject]@{
            Server = $server
            Error = "Inaccessible"
        }
    }
}

$RDPAccess | Export-Excel -Path $ExcelFile -WorksheetName "RDP_Access" -AutoSize

# --- 8) Extraction : Comptes locaux admin pour chaque serveur ---------
$LocalAdmins = foreach ($server in $Servers.Name) {
    try {
        Get-LocalGroupMember -Group "Administrators" -ComputerName $server |
        Select @{N="Server";E={$server}}, Name, ObjectClass
    }
    catch {
        [PSCustomObject]@{
            Server = $server
            Name = "N/A"
            ObjectClass = "Server inaccessible"
        }
    }
}

$LocalAdmins | Export-Excel -Path $ExcelFile -WorksheetName "Local_Admins" -AutoSize

# --- 9) Comptes inactifs et jamais utilisés ---------------------------
$Inactive = Search-ADAccount -AccountInactive -UsersOnly -TimeSpan 90 |
            Select Name, SamAccountName, LastLogonDate, Enabled

$Inactive | Export-Excel -Path $ExcelFile -WorksheetName "Inactive_90_Days" -AutoSize

$NeverLogged = Get-ADUser -Filter * -Properties LastLogonDate |
               Where-Object { $_.LastLogonDate -eq $null } |
               Select Name, SamAccountName

$NeverLogged | Export-Excel -Path $ExcelFile -WorksheetName "Never_Logged" -AutoSize

# --- 10) Fin du script -------------------------------------------------
Write-Host "=============================================="
Write-Host " INVENTAIRE COMPLET GÉNÉRÉ : $ExcelFile"
Write-Host "=============================================="