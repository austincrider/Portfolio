<#
.SYNOPSIS
    Search for and delete an email throughout the M365 organization.
.DESCRIPTION
    Search for and delete an email throughout the M365 organization. Follow the prompts to create a new Compliance Search.
    This program was built to use credential-based authentication. In a future release, the MS Graph API will be leveraged.
.EXAMPLE
    searchAndDeleteEmail.ps1
.NOTES
    =========================================================================== 
     Created by: Austin Crider, https://www.austincrider.com
     Created on: 2022/09/22
     Filename: emailPurge.ps1 
     Version: 1.0
    =========================================================================== 
.LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/new-compliancesearch?view=exchange-ps
#>
############################################## INITIALIZE VARIABLES
function Show-EPLogo {
    Clear-Host
    Write-Host @'
    ============================================

      EEEEE    MM  MM    AAAAAAA  IIIIII  LL
      E       M  M M M   AA   AA    II    LL
      EEEEE  MM   M  MM  AAAAAAA    II    LL
      EE     MM      MM  AA   AA    II    LL
      EE     MM      MM  AA   AA    II    LL
      EEEEE  MM      MM  AA   AA  IIIIII  LLLLL

      PPPPP   UU     UU  RRRRR   GGGGGG  EEEEE
      P    P  UU     UU  R    R  G    G  E
      P    P  UU     UU  R   R   G       EEEEE
      PPPPP   UU     UU  RRRR    G  GGG  EE
      PP      UU     UU  R  R    G    G  EE     
      PP      UUUUUUUUU  R    R  GGGGGG  EEEEE

    =========================================== v1.0
'@
}
function Start-Program {
    Write-Host @'

To search for and delete an email, some information about the message will need to be gathered.

        Press I to install and/or update the Exchange Online module.
                    (Requires you to run this program as Administrator)


        Press S to start.
        Press Q to quit.

'@
}
############################################## FUNCTION - OPEN CONNECTIONS
# Connect to Security & Compliance Center
function Open-Connections {
    Import-Module ExchangeOnlineManagement
    Connect-IPPSSession
    return $true
}
############################################## FUNCTION - INSTALL/UPDATE MODULE
function Check-ExoModule {
    $module = Get-InstalledModule -Name ExchangeOnlineManagement | Format-List Name | Out-String
    $moduleExists = $module.Contains("ExchangeOnlineManagement")
    if ($moduleExists -eq $true){
        return $true
    }
    else {
        Write-Host "The ExchangeOnline module is not installed."
        Read-Host "Press Enter to continue"
        return $false
    }
}
function Get-ExoModule {
    $moduleCheck = Check-ExoModule
    if ($moduleCheck -eq $true){
        Write-Host "You have the module installed."
        $updateChoice = Read-Host "Would you like to update the module? (Y/N)"
        if ($updateChoice -eq 'Y' -Or $updateChoice -eq 'y'){
            Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
        }
        else {
            return $false
        }
    }
    else {
        Read-Host "Would you like to install the module? (Y/N)"
        Write-Host "Attempting to install the module..."
        $installModule = Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
        Write-Host $installModule
        Read-Host "Press Enter to continue"
        
    }
}
############################################## FUNCTION - DELETE OLD SEARCH
function Remove-OldSearch {
    Remove-ComplianceSearch -Identity "PhishSearch1" -Confirm:$false
}
############################################## FUNCTION - CREATE NEW SEARCH
function Start-NewSearchCreation {
    # Prompt for search variables
    Write-Host @'
    ------------------------------------
    Please input the search parameters...

'@
    $emailSubject = Read-Host -prompt "Word(s) from the email subject (example - Student Service)"
    $emailSentDate = Read-Host -prompt "Input the date the email was sent (example - MM/DD/YYYY)"
    $emailSentQuery = 'sent:' + $emailSentDate
    $emailSubjectQuery = '"*' + $emailSubject + '*"'
    Write-Host "Date: " $emailSentQuery
    Write-Host "Subject: " $emailSubjectQuery
    # Creates the compliance search using input variables
    New-ComplianceSearch `
    -Name PhishSearch1 `
    -ExchangeLocation All `
    -Description 'Search and Delete Script' `
    -ContentMatchQuery "($emailSubjectQuery)($emailSentQuery)"
}
############################################## FUNCTION - BEGIN SEARCH
function Start-EmailSearch {
    Start-ComplianceSearch -Identity PhishSearch1
    $searchStatus = $false
    while ($searchStatus -eq $false) {
        # Determine status and end loop when "Complete"
        $searchStatus = Get-SearchStatus
    }
    return $true
}

############################################## FUNCTION - GET SEARCH STATUS
function Get-SearchStatus {
    while ($status -ne $true) {
        Show-EPLogo
        $statusString = Get-ComplianceSearch -Identity PhishSearch1 | Format-List Status | Out-String
        $status = $statusString.Contains("Completed")
        Write-Host "Search in progress. This could take a while." -NoNewLine -ForegroundColor black -BackgroundColor green
        Start-Sleep -Seconds 5
        Write-Host "." -NoNewLine -ForegroundColor black -BackgroundColor green
        Start-Sleep -Seconds 5
        Write-Host "." -NoNewLine -ForegroundColor black -BackgroundColor green
        Start-Sleep -Seconds 5
        Write-Host "." -NoNewLine -ForegroundColor black -BackgroundColor green
        Start-Sleep -Seconds 5
        Write-Host "." -NoNewLine -ForegroundColor black -BackgroundColor green
        Start-Sleep -Seconds 5
        Write-Host "." -ForegroundColor black -BackgroundColor green
        Start-Sleep -Seconds 5
    }
    return $true
}
############################################## FUNCTION - CREATE PREVIEW
function Start-EmailPreview {
    # Creates a preview action. This appends _Preview to the SearchName:
    New-ComplianceSearchAction -SearchName PhishSearch1 -Preview
    $previewStatus = $false
    while ($previewStatus -eq $false) {
        # Determine status and end loop when "Complete"
        $previewStatus = Get-PreviewStatus
    }
    Read-Host "Preview creation complete. Press Enter to view the results"
    $results = (Get-ComplianceSearchAction PhishSearch1_Preview | Select-Object -ExpandProperty Results) -split ","
    Write-Host $results
    Read-Host "Press Enter to continue"
    return $true
}
############################################## FUNCTION - GET PREVIEW STATUS
function Get-PreviewStatus {
    while ($status -ne $true) {
        Show-EPLogo
        $statusString = Get-ComplianceSearchAction -Identity PhishSearch1_Preview | Format-List Status | Out-String
        $status = $statusString.Contains("Completed")
        Write-Host "Preview is generating. This could take a while." -NoNewLine -ForegroundColor black -BackgroundColor yellow
        Start-Sleep -Seconds 5
        Write-Host "." -NoNewLine -ForegroundColor black -BackgroundColor yellow
        Start-Sleep -Seconds 5
        Write-Host "." -NoNewLine -ForegroundColor black -BackgroundColor yellow
        Start-Sleep -Seconds 5
        Write-Host "." -ForegroundColor black -BackgroundColor yellow
        Start-Sleep -Seconds 5
    }
    return $true
}
############################################## FUNCTION - BEGIN PURGE
function Start-EmailPurge {
    # Creates a purge job. This appends _Purge to the SearchName and starts the purge:
    New-ComplianceSearchAction -SearchName PhishSearch1 -Purge -PurgeType SoftDelete -Confirm:$false
    $purgeStatus = $false
    while ($purgeStatus -eq $false) {
        # Determine status and end loop when "Complete"
        $purgeStatus = Get-PurgeStatus
    }
    return $true
}
############################################## FUNCTION - GET PURGE STATUS
function Get-PurgeStatus {
    while ($status -ne $true) {
        Show-EPLogo
        $statusString = Get-ComplianceSearchAction -Identity PhishSearch1_Purge | Format-List Status | Out-String
        $status = $statusString.Contains("Completed")
        Write-Host "Performing the purge. This could take a while." -NoNewLine -ForegroundColor white -BackgroundColor DarkRed
        Start-Sleep -Seconds 5
        Write-Host "." -NoNewLine -ForegroundColor white -BackgroundColor DarkRed
        Start-Sleep -Seconds 5
        Write-Host "." -NoNewLine -ForegroundColor white -BackgroundColor DarkRed
        Start-Sleep -Seconds 5
        Write-Host "." -ForegroundColor white -BackgroundColor DarkRed
        Start-Sleep -Seconds 5
    }
    return $true
}
############################################## FUNCTION - CLOSE CONNECTIONS
function Close-Connections {
    Disconnect-ExchangeOnline -Confirm:$false
}
############################################## BEGIN PROGRAM
$start = $true
while ($start -ne $false) {
    Show-EPLogo
    Start-Program
    $startSelection = Read-Host "Selection"
    if ($startSelection -eq 'S' -Or $startSelection -eq 's'){
        Open-Connections
        Write-Host "Attempting to remove old searches... "
        Remove-OldSearch
        Start-NewSearchCreation # Prompts for search variables
        $searchChoice = Read-Host "Ready to begin search? (Y/N)"
        if ($searchChoice -eq 'Y' -Or $searchChoice -eq 'y') {
            $searchFunction = Start-EmailSearch
            if ($searchFunction -eq $true) {
                Show-EPLogo
                Write-Host "The search is complete." -ForegroundColor black -BackgroundColor green
                $seePreview = Read-Host "Would you like to see a preview of the results? (Y/N)"
                if ($seePreview -eq 'Y'){
                    $previewFunction = Start-EmailPreview
                    if ($previewFunction -eq $true) {
                        Write-Host @'
                        --------------------------------
                        The Preview process is complete.
                        
                        This is the final prompt before beginning the purge.
                        Proceed with caution.

'@
                        $purgeChoice = Read-Host "Would you like to purge the emails? (Y/N)"
                        if ($purgeChoice -eq 'Y' -Or 'y'){
                            $purgeFunction = Start-EmailPurge
                            if ($purgeFunction -eq $true){
                                Write-Host "The email purge process is complete."
                                Read-Host "Press Enter"
                                Remove-OldSearch
                                }
                            }
                            
                    }
                    else {
                        Write-Host @'
                    This is the final prompt before beginning the purge.
                    Proceed with caution.

'@
                        $purgeChoice = Read-Host "Would you like to begin the purge? (Y/N)"
                        if ($purgeChoice -eq 'Y' -Or 'y'){
                            $purgeFunction = Start-EmailPurge
                            if ($purgeFunction -eq $true){
                                Write-Host "The email purge process is complete."
                                Read-Host "Press Enter"
                            }
                            }
                        }
                    }
            else {
                $noSearch = Read-Host "Would you like to quit the program? (Y/N)"
                if ($noSearch -eq 'Y' -Or $noSearch -eq 'y') {
                    Remove-OldSearch
                    Close-Connections
                    $start = $false
                }
            }
            }
        else {
            $noSearch = Read-Host "Would you like to quit the program? (Y/N)"
            if ($noSearch -eq 'Y' -Or $noSearch -eq 'y') {
                Remove-OldSearch
                Close-Connections
                $start = $false
                }
        }
        }
    }
    elseif ($startSelection -eq 'I' -Or $startSelection -eq 'i') {
        Write-Host @'

        This process will attempt to install the ExchangeOnlineManagement module.

                If this process does not succeeed, please follow these steps to install the module manually:

                1. Open a PowerShell as administrator
                2. Run the following command: Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
                
'@
        Read-Host "Press Enter to continue to the installation process"
        Get-ExoModule
    }
    elseif ($startSelection -eq 'Q' -Or $startSelection -eq 'q') {
            Write-Host "Goodbye."
            $start = $false
    }
    else {
        Write-Host "ERROR: Something is not right. You input `"$startSelection`"."  -ForegroundColor black -BackgroundColor red
        Read-Host "Press Enter to start over"
    }
}
Close-Connections