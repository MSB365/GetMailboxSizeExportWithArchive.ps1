<#
.SYNOPSIS
GetMailboxSizeExportWithArchive.ps1

.DESCRIPTION 
This powershell script can be used to generate a Certificate Signing Request (CSR) using the SHA256 signature algorithm and a 2048 bit key size (RSA). Subject Alternative Names are supported.


.PARAMETER



.EXAMPLE
.\GetMailboxSizeExportWithArchive.ps1 


.NOTES
Written by: Drago Petrovic
 



Find me on:

* LinkedIn:	https://www.linkedin.com/in/drago-petrovic-86075730/
* Xing:     https://www.xing.com/profile/Drago_Petrovic
* Website:  https://blog.abstergo.ch
* GitHub:   https://github.com/MSB365


Change Log:
v1.0
- initial version


--- keep it simple, but significant ---

.COPYRIGHT
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>
Write-Host "keep it simple but significant" -ForegroundColor green
Write-Host "Enter the Certificate informations below" -ForegroundColor cyan

# Set Variable
$OrganizationalUnit = Read-Host "Enter OrganizationalUnit e.g. contoso.lan/CUSTOMERS/Employees/Standard"
$OutputPath = Read-Host "Enter the Output Path e.g. C:\Temp\ "

# Script
$Mailboxes = Get-Mailbox -ResultSize unlimited | where {$_.OrganizationalUnit -eq $OrganizationalUnit} | foreach { 

write-host $_.alias
    $user = ""
    $temp = Get-MailboxStatistics -Identity $_.alias
    $user = $user + $temp.DisplayName
    Write-Host $temp.DisplayName
    $user = $user + ";" + $temp.TotalItemSize
    Write-Host $temp.TotalItemSize


    if($_.ArchiveGuid -ne "00000000-0000-0000-0000-000000000000"){    
        $temp1 = Get-MailboxStatistics -Identity $_.alias -Archive
        Write-Host $temp1.TotalItemSize
        $user = $user + ";" + $temp1.TotalItemSize
      }
      else {
        $user = $user + ";;"
      }

    
    Write-host $user

  

$user >> "$OutputPath\Output-Statistic.txt"
}