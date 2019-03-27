# ContactSplitter.ps1
# Takes a file containing a Gmail dump of your contacts and splits it into individual files that can be imported into Microsoft Outlook
# 27-MAR-2019: GGARIEPY: Version 1.00 Creation

<#
.SYNOPSIS 
Creates individual .vcf contact files from a .vcf file containing multiple contacts 

.DESCRIPTION
Google can provide Gmail users a "dump" of their contacts, all in one large .VCF file.

Unfortunately, Outlook does not appear to be able to import more than the first contact from this list.

This script solves that problem by parsing the Google-supplied file and creating individual .VCF files for each contact.
    
.PARAMETER BulkFileName
-BulkFileName the name of the file to be split up (mandatory)

.PARAMETER TargetDir
-TargetDir tells the script where to write the individual .VCF files

.INPUTS
None. You cannot pipe objects to ContactSplitter.ps1

.OUTPUTS
ContactSplitter.ps1 generates console messages about the files being written as it progresses

.EXAMPLE
C:\git\ContactSplitter> .\ContactSplitter.ps1 -BulkFilename C:\Users\username\Downloads\ContactBackup.vcf -TargetDir c:\Users\username\Documents\Contacts'

#>


[CmdletBinding()]

Param(
    [Parameter(Mandatory=$true)]
    [string] 
    $BulkFileName,

    [Parameter(Mandatory=$false)]
    [string] 
    $TargetDir = "$env:HOMEDRIVE" + "$env:HOMEPATH\Documents\Contacts"
)

Clear-Host
$ErrorActionPreference = 'Stop'
$contactdata = Get-Content -Path $BulkFileName 
$contactname = '';
$newcontact= @()
$incontact = $false
foreach ($line in $contactdata){

    if($line -eq 'BEGIN:VCARD') {
        $incontact = $true       
    }
    elseif ($line -eq 'END:VCARD') {
        $incontact = $false
    }
    elseif ($line -like 'FN:*' -or ($contactname.Length -eq 0 -and $line -like 'ORG:*'))
    {
        $contactnamearr = $line -split ':'
        $contactname = $contactnamearr[1]
    }

    $newcontact += $line

    if ($incontact -eq $false -and $contactname.Length -gt 0){
        $contactname = $contactname.Replace('\', '-') 
        $contactname = $contactname.Replace('/', '-') 
        $ContactPath = Join-Path -Path $TargetDir -ChildPath "$contactname.vcf" 
        $ContactPath = $ContactPath.Replace('..', '.')
        $ContactPath = $ContactPath.Replace('  ', ' ')
        $newcontact | Out-File $ContactPath -Force -ErrorAction Stop
        Write-Host "Created $ContactPath"
        $newcontact =  @()
        $contactname = ''
    }
    elseif ($incontact -eq $false) {
        Write-Host "Incomplete contact, discarding:`n$newcontact"
        $newcontact =  @()
        $contactname = ''
    }
}