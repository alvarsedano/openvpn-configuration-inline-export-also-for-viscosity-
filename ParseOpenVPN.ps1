####
### Exporting openvpn client to inline unique file
####
# This ps1 script exports multifile openvpn client configs
# to one inline .ovpn file
#
# The Viscosity (also Tunnelblick) format is valid for MacOS
# This file format uses the UTF8bom encoding.
# 
# openVPN Client must be installed in this computer.

# Tested on PowerShell 5.1
# Created by Alvaro Sedano Galindo. al_sedano@hotmail.com

$ErrorActionPreference = 'SilentlyContinue'

#
# Functions
#

Function Get-BeginEnd {
    Param([Parameter(Mandatory=$true, Position=0)]
          [string]$path)

    [string[]]$text = Get-Content $path -Encoding UTF8
    [int]$beg = ($text | Select-String -Pattern "-BEGIN " -Encoding utf8 | Select -ExpandProperty LineNumber) - 1
    [int]$end =  $text | Select-String -Pattern "-END "   -Encoding utf8 | Select -ExpandProperty LineNumber
    $text[$beg..$end]
}

Function BufferAdd {
    Param([Parameter(Mandatory=$true, Position=0)]
          [ref]$dir,
          [Parameter(Mandatory=$true, Position=1)]
          [ref]$buffer,
          [Parameter(Mandatory=$true, Position=2)]
          [string]$pa)

    $buffer.Value += "<$pa>"
    $buffer.Value += Get-BeginEnd -path ($dir.value)
    $buffer.Value += "</$pa>"
}

Function IfExists {
    Param([Parameter(Mandatory=$true, Position=0)]
          [ref]$ovpn,
          [Parameter(Mandatory=$true, Position=1)]
          [ref]$dir,
          [Parameter(Mandatory=$true, Position=2)]
          [ref]$buffer,
          [Parameter(Mandatory=$true, Position=3)]
          [string]$pa)

    $pattern = "^$pa (?<a>.*) *$"
    [string]$a = ($ovpn.Value) -match $pattern
    if ($a -ne $null -and $a -match $pattern) {
        # Pattern found
        [string]$arch = "$($dir.Value)\$($Matches.a)"

        if(Test-Path $arch) {
            BufferAdd -dir ([ref]$arch) -buffer $buffer -pa $pa
        }
        else {
            Write-Host "ERROR: The $pa file '$arch' was not found. Process stopped." -BackgroundColor DarkRed
            Exit(2)
        }
    }
}

#
# BODY
#

# Get from registry the OpenVPN Client path
[string]$rutaREG = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\OpenVPN"
if (-not (Test-Path($rutaREG))) {
    Write-Output "No openvpn installation found. You must install openvpn client to use this script." -BackgroundColor DarkRed
    Exit (1)
}

# openVPN Client installation found
[string[]]$rutaConfig = (Get-ItemProperty -Path $rutaREG).config_dir
$rutaConfig += @("$env:USERPROFILE\OpenVPN\config" `
                , "$([Environment]::GetFolderPath("MyDocuments"))" )

[string]$openSSL = ((Get-ItemProperty -Path $rutaREG).exe_path).Replace("openvpn.exe", "openssl.exe")

# List every ovpn tunnel
[string[]]$tunnels = (Get-ChildItem -Filter "*.ovpn" -Path $rutaConfig -Depth 1 -Recurse).Fullname

if ($tunnels -eq $null -or $tunnels.Count -eq 0) {
    Write-Host "No ovpn config tunnels found in" -BackgroundColor DarkRed
    foreach ($p in $rutaConfig) {
        Write-Host "`t'$p'"
    }
    Write-Host "and its level 1 subfolders." -BackgroundColor DarkRed
    exit(3)
}

for($i=0; $i -lt $tunnels.Count; $i++) {
    Write-Host $([string]::Format("[{0}] - {1}", $i+1, $tunnels[$i]) )
}

[int]$max = $tunnels.Count
do {
    [string]$resp = Read-Host "`nWhich ovpn tunnel do you want to export? [1-$max] (Ctrl+C to exit)"
} While ([int]$resp -lt 1 -or [int]$resp -gt $max)
[string]$tun = $tunnels[$resp-1]
[string[]]$content = Get-Content -Path $tun
[string]$retls = "^tls-auth (?<a>.*) (?<b>[0-1])"
[string]$reca = "^ca (?<a>.*) *$"
[string]$recert = "^cert (?<a>.*) *$"
[string]$rekey = "^key (?<a>.*) *$"
[string]$repkcs = "^pkcs12 (?<a>.*)"

[string[]]$exported = @()
$exported += $content | Select-String -NotMatch -Pattern $retls  | `
                            Select-String -NotMatch -Pattern $reca   | `
                            Select-String -NotMatch -Pattern $recert | `
                            Select-String -NotMatch -Pattern $rekey  | `
                            Select-String -NotMatch -Pattern $repkcs | `
                            Select-String -NotMatch -Pattern '^#'    | `
                            Select-String -NotMatch -Pattern '^;'
#$exported += ""

[System.IO.FileSystemInfo]$ar = Get-ChildItem -Path ($tunnels[$resp-1])
[string]$dir = $ar.Directory.FullName

#PKCS12 section found
[string]$a = $content -match $repkcs
if ($a -ne $null -and $a -match $repkcs) {
    # pkcs12 section found
    [string]$arch = "$dir\$($Matches.a)"

    #Extract from .p12 the ca,crt,key files
    $f1CA  = New-TemporaryFile
    $f2cer = New-TemporaryFile
    $f3key = New-TemporaryFile
    try {
        & "$openSSL" pkcs12 -in "$arch" -nomac -nokeys -cacerts -out "$($f1CA.FullName)"  -passin pass:
        & "$openSSL" pkcs12 -in "$arch" -nomac -nokeys -clcerts -out "$($f2cer.FullName)" -passin pass:
        & "$openSSL" pkcs12 -in "$arch" -nomac -nodes  -nocerts -out "$($f3key.FullName)" -passin pass:
    
        BufferAdd -dir ([ref]($f1CA.FullName))  -buffer ([ref]$exported) -pa 'ca'
        BufferAdd -dir ([ref]($f2cer.FullName)) -buffer ([ref]$exported) -pa 'cert'
        BufferAdd -dir ([ref]($f3key.FullName)) -buffer ([ref]$exported) -pa 'key'
    }
    catch {}
    finally {
        Remove-Item $f1CA.FullName  -Force
        Remove-Item $f2cer.FullName -Force
        Remove-Item $f3key.FullName -Force
    }
}
else {
    # No PKCS12 section, search for ca, cert, key section
    IfExists -ovpn ([ref]$content) -dir ([ref]$dir) -buffer ([ref]$exported) -pa 'ca'
    IfExists -ovpn ([ref]$content) -dir ([ref]$dir) -buffer ([ref]$exported) -pa 'cert'
    IfExists -ovpn ([ref]$content) -dir ([ref]$dir) -buffer ([ref]$exported) -pa 'key'
}

#TLS-AUTH section
[string]$a = $content -match $retls
if ($a -ne $null -and $a -match $retls) {
    # tls-auth section found
    [string]$arch = "$dir\$($Matches.a)"
    [string]$keydir = "key-direction $($Matches.b)"

    [string[]]$c = Get-Content $arch

    $exported += $keydir
    $exported += "<tls-auth>"
    $exported += Get-Content $arch
    $exported += "</tls-auth>"
}

# Obtener nombre archivo
[int]$beg = 1 + $tun.LastIndexOf('\')
[string]$fSal = $tun.Substring($beg)

# To choose the exported file type (Viscosity uses UTF8bom encoding), Default/ANSI encodig for the rest
[string]$isVisc = Read-Host "Export for MacOS (Viscosity/Tunnelblick) (UTF8 w BOM)? [y/N]"
if ($isVisc -eq 'y') {
    # Viscosity. Export with "UTF8 w BOM" encoding
    [string]$nombre = $fSal.Replace('.ovpn', '')
    [string[]]$visco = @('#-- Config Generated by ParseOpenVPN for Viscosity --#', '' `
                        ,'#viscosity startonopen false' `
                        ,'#viscosity dhcp true' `
                        ,'#viscosity dnssupport true' `
                        ,"#viscosity name $nombre")
    
    $UTF8bomEnc = [System.Text.UTF8Encoding]::new($true);
    $fSal = $fSal.Replace('.ovpn', '-viscosity.ovpn')
    [System.IO.File]::WriteAllLines("$env:UserProfile\$fSal", $visco + $exported, $UTF8bomEnc)
}
else {
    # Rest of platforms. Export with ANSI encoding
    # TODO: validation pending: is 'default' valid for every Windows OS languages?
    $fSal = $fSal.Replace('.ovpn', '-inline.ovpn')
    $exported | Out-File -Encoding default -FilePath "$env:UserProfile\$fSal"
}

Write-host $([string]::Format("`nInline config file '{0}' created in folder '{1}'.", $fSal, $env:UserProfile)) -BackgroundColor DarkGreen
