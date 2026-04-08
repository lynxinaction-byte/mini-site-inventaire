# Script d'inventaire automatisé pour le parc informatique
# VERSION : 
# 1.4 - 31/03/2026 : Correction des chemins (Rangement automatique)
# 1.3 - 31/03/2026 : Ajout de l'export CSV pour Excel/EasyVista (Sami)

# --- CONFIGURATION ---
$ServeurCollecte = "SERVEUR_HOST"
$Partage = "Inventaire$"
$SambaPath = "\\$ServeurCollecte\$Partage"

# === DEFINITION DES CHEMINS ===
$FolderTXT = Join-Path $SambaPath "Rapports_Detailles"
$FolderCSV = Join-Path $SambaPath "Base_Donnees"

# Création des dossiers uniquement si le partage existe
if (Test-Path $SambaPath) {
    foreach ($folder in @($FolderTXT, $FolderCSV)) {
        if ($folder -and -not (Test-Path $folder)) {
            New-Item -ItemType Directory -Path $folder | Out-Null
        }
    }
} else {
    Write-Host "[!] Partage inaccessible : $SambaPath" -ForegroundColor Yellow
}

# === FICHIERS ===
$Timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$FileName  = "INV_$($env:COMPUTERNAME)_$($env:USERNAME)_$Timestamp.txt"
$CSVName   = "Inventaire_Global.csv"

$FinalPathTXT = Join-Path $FolderTXT $FileName
$FinalPathCSV = Join-Path $FolderCSV $CSVName

Write-Host "Génération du rapport d'inventaire..." -ForegroundColor Cyan

# --- COLLECTE DES INFOS SYSTEME ---
$osObj = Get-CimInstance Win32_OperatingSystem
$OSFriendly = "$($osObj.Caption) $($osObj.OSArchitecture)"
$OSBuild = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentBuild

# Dernier boot
$BootRaw = $osObj.LastBootUpTime
$LastBoot = $BootRaw.ToString("dd/MM/yyyy HH:mm")

# Dernier logon 
try {
    $LastLogonEvent = Get-WinEvent -FilterHashtable @{
        LogName='Security'; Id=4624
    } -MaxEvents 1 -ErrorAction SilentlyContinue

    if ($LastLogonEvent) {
        $LastLogon = $LastLogonEvent.TimeCreated.ToString("dd/MM/yyyy HH:mm")
    } else {
        $LastLogon = (Get-Item "C:\Users\$env:USERNAME").LastWriteTime.ToString("dd/MM/yyyy HH:mm")
    }
} catch {
    $LastLogon = "Non disponible"
}

# --- AUTRES VARIABLES ---
$DateFix = Get-Date -Format "dd/MM/yyyy HH:mm"
$UserSession = $env:USERNAME
$PCName = $env:COMPUTERNAME

# Email utilisateur
$UserEmail = "Non disponible"
try {
    $searcher = [adsisearcher]"(samAccountName=$env:USERNAME)"
    $result = $searcher.FindOne()
    if ($result -ne $null) { $UserEmail = [string]$result.Properties.mail }
} catch { 
    $UserEmail = "Hors Domaine / Erreur AD" 
}

# Hardware
$serial = (Get-CimInstance Win32_BIOS).SerialNumber
$cpu = (Get-CimInstance Win32_Processor).Name.Trim()
$ram = [math]::round((Get-CimInstance Win32_PhysicalMemory | Measure-Object Capacity -Sum).Sum / 1GB, 0)
$IPs = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike "127*" }).IPAddress -join " | "

# --- LOGICIELS ---
$Paths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$rawApps = Get-ItemProperty $Paths -ErrorAction SilentlyContinue |
           Where-Object { $_.DisplayName -ne $null -and $_.SystemComponent -ne 1 }

$appsList = $rawApps |
    ForEach-Object {
        # Nettoyage du nom : suppression version + édition
        $clean = $_.DisplayName `
            -replace '\s*\(x64 edition\)', '' `
            -replace '\s*\(x86\)', '' `
            -replace '\s*x64', '' `
            -replace '\s*x86', '' `
            -replace '\s*\d+(\.\d+)+', '' `
            -replace '\s+$',''

        $_ | Add-Member -NotePropertyName CleanName -NotePropertyValue $clean -Force
        $_
    } |
    Group-Object CleanName |
    ForEach-Object {
        # On trie par version et on garde la plus récente
        $latest = $_.Group | Sort-Object {
            try { [version]$_.DisplayVersion } catch { [version]"0.0.0.0" }
        } -Descending | Select-Object -First 1

        "$($latest.CleanName) [$($latest.DisplayVersion)]"
    } | Sort-Object

$AppsText = $appsList -join "`r`n"

# --- CONSTRUCTION DU RAPPORT TEXTE ---
$Report = @"
[LOG]
Date_Collecte   : $DateFix
Fichier_Source  : $FileName

[IDENTITE]
User            : $UserSession
Email           : $UserEmail
PC              : $PCName
S/N             : $serial
Last_Logon      : $LastLogon

[SYSTEME]
OS              : $OSFriendly
Build_OS        : $OSBuild
Dernier_Boot    : $LastBoot
IP              : $IPs

[HARDWARE]
CPU             : $cpu
RAM             : $ram GB

[LOGICIELS]
$AppsText
"@

# --- CREATION DE LA LIGNE CSV ---
$LigneCSV = [PSCustomObject]@{
    Date        = $DateFix
    Utilisateur = $UserSession
    PC          = $PCName
    Serial      = $serial
    OS          = $OSFriendly
    Build       = $OSBuild
    Boot        = $LastBoot
    RAM         = $ram
    IP          = $IPs
}

# --- EXPORT RÉSEAU ---
if (Test-Path $SambaPath) {
    try {
        $Report | Out-File -FilePath $FinalPathTXT -Encoding utf8 -ErrorAction Stop
        $LigneCSV | Export-Csv -Path $FinalPathCSV -Append -NoTypeInformation -Delimiter ";" -Encoding UTF8
        Write-Host "[+] Inventaire et Base de données mis à jour !" -ForegroundColor Green
    } catch {
        Write-Host "[!] Erreur : $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "[!] Partage inaccessible, export impossible." -ForegroundColor Yellow
}

Start-Sleep -Seconds 3
