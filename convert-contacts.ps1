<#
    .SYNOPSIS
        Converts Nokia CSV contacts file (English only) to a simple vCard file.

    .DESCRIPTION
        Converts Nokia CSV contacts file (English only) to a simple vCard file.
        Takes only few fields.

    .PARAMETER FileName
        Filename of CSV

    .PARAMETER ics
        Will additionally generate vCalendar file.

    .EXAMPLE
        C:\PS> ConvertTo-VCF.ps1 contacts.csv

        Will create my_contacts.vcf file in the same folder there it running.

    .EXAMPLE
        C:\PS> convert-contacts my_contacts.csv -ics

        Will create my_contacts.ics file in the same folder there it running.

    .NOTES
        Author: Alexey Matashkin
        Date:   August 8, 2013
#>

param(
    [Parameter(Mandatory=$true,Position=0,HelpMessage="Enter name of CSV file")]
    [String]$Filename,
    [switch]$ics
)

# if (!($Filename)) {echo "No file!";exit 3}

if (!(Test-Path($Filename))) {
    "No such file. Check filename."
    exit 1000
}

$File = Get-Item $FileName
$FileVCF = $File.Basename + ".vcf"
$FileICS = $File.Basename + ".ics"

if (Test-Path($FileVCF)) {
    "Can't write, file " + $FileVCF + " already exist!"
    exit 1001
}

if ($ics) {
    if (Test-Path($FileICS)) {
        "Can't write, file " + $FileICS + " already exist!"
        exit 1002
    }

}

$contacts = ConvertFrom-Csv (Get-Content $File.FullName).Replace(';',',')
$vcard = ""
foreach ($contact in $contacts) {
    $vcard = $vcard + "BEGIN:VCARD" + "`r`n"
    $vcard = $vcard + "VERSION:3.0" + "`r`n"
    $vcard = $vcard + "N:" + $contact."First name" + ";" + $contact."Last Name" + "`r`n"
    $vcard = $vcard + "FN:" + $contact."First name" + " " + $contact."Last Name" + "`r`n"
    $vcard = $vcard + "TEL:" + $contact."General phone" + "`r`n"
    $vcard = $vcard + "EMAIL:" + $contact."General email" + "`r`n"
    $vcard = $vcard + "END:VCARD" + "`r`n"
}

$vcard | Add-Content ($FileVCF) 
write-host ("Generated " + $FileVCF)


if ($ics) {

    # Generate file contents
    $b64vcard = ConvertTo-Base64 -Path $FileVCF -NoLineBreak
    $b64split = " " + ($b64vcard -replace "(\w{74})","`$1`r`n ")
    
    $datestart = Get-Date -uformat %Y%m%dT%H%M00
    $dateend = Get-Date -uformat %Y%m%dT%H%M01
    $vcal = ""

    $vcal = $vcal + "BEGIN:VCALENDAR" + "`r`n"
    $vcal = $vcal + "VERSION:2.0" + "`r`n"
    $vcal = $vcal + "BEGIN:VEVENT" + "`r`n"
    $vcal = $vcal + "DTSTART;TZID=Europe/Moscow:$datestart" + "`r`n"
    $vcal = $vcal + "DTEND;TZID=Europe/Moscow:$dateend" + "`r`n"
    $vcal = $vcal + "SUMMARY:iPhone Contact" + "`r`n"
    $vcal = $vcal + "DTSTAMP:$datestart" + "Z" + "`r`n"
    $vcal = $vcal + "ATTACH;VALUE=BINARY;ENCODING=BASE64;FMTTYPE=text/directory;" + "`r`n"
    $vcal = $vcal + " X-APPLE-FILENAME=iphonecontact.vcf:" + "`r`n"
    # $vcal = $vcal + ($b64vcard -replace "(\w{74})"," `$1`r`n ") + "`r`n"
    $vcal = $vcal + $b64split + "`r`n"
    $vcal = $vcal + "END:VEVENT" + "`r`n"
    $vcal = $vcal + "END:VCALENDAR" + "`r`n"

    $vcal | Add-Content ($FileICS)
    write-host ("Generated " + $FileICS)
    
}
