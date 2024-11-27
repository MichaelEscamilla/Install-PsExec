param (
    [ValidateSet("x64", "x86")]
    [string]$Architecture = "x64",
    [switch]$AcceptEULA,
    [switch]$RunIEasSYSTEM
)

### Check if connected to the internet, No Internet, No PsExec
try {
    Write-Host "Checking internet connection..."
    $null = Test-Connection -ComputerName google.com -Count 1 -ErrorAction Stop
    Write-Host "Connected to the internet"
}
catch {
    Write-Error "Not connected to the internet"
    exit 1
}

### Set Download URL based on architecture
if ($Architecture -eq "x64") {
    $DownloadURL = "https://live.sysinternals.com/PsExec64.exe"
}
else {
    $DownloadURL = "https://live.sysinternals.com/PsExec.exe"
}
    Write-Host "Download URL: [$DownloadURL]"

### Parse File Name from Download URL
$FileName = $DownloadURL.Split("/")[-1]

### Set Save Path
# This method will grab the OneDrive folder if Backup is enabled
$SavePath = [Environment]::GetFolderPath('MyDocuments')

### Combine Save Path and File Name
$SaveFile = Join-Path -Path $SavePath -ChildPath $FileName

### Download PsExec
try {
    Write-Host "Downloading PsExec to: [$SaveFile]"
    Invoke-WebRequest -Uri $DownloadURL -OutFile $SaveFile -ErrorAction Stop
    Write-Host "Download Successful"
}
catch {
    Write-Error "Failed to download PsExec from $DownloadURL"
    exit 2
}

### Create Accept EULA Registry Key if AcceptEULA switch is set
if ($AcceptEULA) {
    # Define the registry path for ZoomIt settings
    $RegPath = "HKCU:\Software\Sysinternals\PsExec"
    
    # Define the registry entry name for EULA acceptance
    $RegName = "EulaAccepted"
    
    # Define the registry entry value to indicate EULA acceptance
    $RegValue = "1"
    
    # Create the registry path if it doesn't exist
    if (-not (Test-Path $RegPath)) {
        New-Item -Path $RegPath -Force | Out-Null
    }
    
    # Create or update the registry entry to accept the EULA
    New-ItemProperty -Path $RegPath -Name $RegName -Value $RegValue -PropertyType DWord -Force | Out-Null
}

### Launch PsExec and Run iexplorer.exe as SYSTEM
if ($RunIEasSYSTEM) {
    Write-Host "'Run Internet Explorer' Option selected"
    # Get the Security Principal
    $Global:CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())

    # Check if the script is running as an administrator
    Write-Host "Checking if script is running as an administrator"
    if (($CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
        Write-Host "Script is running as an administrator"
        Start-Process -FilePath $SaveFile -ArgumentList "-i -s `"C:\Program Files\Internet Explorer\iexplore.exe`""
        Write-Host "Successfully launched Internet Explorer as SYSTEM" -ForegroundColor Green
    }
    else {
        Write-Host "Current Principal: [$($CurrentPrincipal)]"
        Write-Host "Current Principal Identity Name: [$($CurrentPrincipal.Identity.Name)]"
        Write-Host "Current Principal IsInRole: [$($CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))]"
        Write-Host "This script must be run as an administrator to start Internet Explorer as SYSTEM" -ForegroundColor Red
        Write-Host "Please rerun the script as an administrator" -ForegroundColor Red
        exit 3
    }
}