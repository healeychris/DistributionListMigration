<#	
    .NOTES
    ===========================================================================
    Created with: 	VS
    Created on:   	28/01/2022
    Created by:   	Chris Healey
    Organization: 	
    Filename:       Recreate-DistributionGroup.ps1
    Project path:   https://github.com/healeychris/DistributionListMigration
    Org Author:     Joe Palarchio (based on Version: 1.0) 
    ===========================================================================
    .DESCRIPTION
    Copies attributes of a synchronized group to a placeholder group and CSV file.  After 
    initial export of group attributes, the on-premises group can have the attribute
    "AdminDescription" set to "Group_NoSync" which will stop it from be synchronized.
    The "-Finalize" switch can then be used to write the addresses to the new group and
    convert the name.  The final group will be a cloud group with the same attributes as
    the previous but with the additional ability of being able to be "self-managed".
    Once the contents of the new group are validated, the on-premises group can be deleted.
    .NOTES

    Run Order - In 365 to create duplicate group of synced group
    .\Recreate-DistributionGroup.ps1 -Group "DL-Marketing" -CreatePlaceHolder

    Run on prem to remove the objects from sync and create contact objects. (not synced to 365)
    .\Recreate-DistributionGroup.ps1 -Group "DL-Marketing" -Contact

    Run to Finalize the cloud group and cut over the original name
    .\Recreate-DistributionGroup.ps1 -Group "DL-Marketing" -Finalize
#>

<#
	.PARAMETER Group
		Name of group to recreate.

	.PARAMETER CreatePlaceHolder
		Create placeholder group.

	.PARAMETER Finalize
		Convert placeholder group to final group.

    .PARAMETER Contact
		Create a Contact based on Group for Onpremise emailing.

    	.EXAMPLE #1
        	.\Recreate-DistributionGroup.ps1 -Group "DL-Marketing" -CreatePlaceHolder

    	.EXAMPLE #2
        	.\Recreate-DistributionGroup.ps1 -Group "DL-Marketing" -Finalize

        .EXAMPLE #3
        .\Recreate-DistributionGroup.ps1 -Group "DL-Marketing" -Contact
#>


Param(
    [Parameter(Mandatory=$True)]
        [string]$Group,

    [Parameter(Mandatory=$False)]
        [switch]$CreatePlaceHolder,

    [Parameter(Mandatory=$False)]
        [switch]$Finalize,

    [Parameter(Mandatory=$False)]
        [switch]$Contact
)

$ContactGroupOU             =       'OU=MovedDistributionGroups,OU=Siemens SIS London,OU=Contacts,OU=Standard,OU=Business,DC=national,DC=core,DC=bbc,DC=co,DC=uk'
$DCServer                   =       'BGB01DC1180'               # DC ServerName
$ExportDirectory            =       ".\ExportedAddresses\"
$FullGroupExportDirectory   =       ".\FullGroupExport\"
$FailedGroups               =       ".\FailedGroups\"
Clear-Host


# FUNCTION - WriteTransaction Log function    
function WriteTransactionsLogs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string]$Task,
 
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateSet('Information','Warning','Error','Completed','Processing')]
        [string]$Result,
 
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [string]$ErrorMessage,
    
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateSet('True','False')]
        [string]$ShowScreenMessage,
 
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [string]$ScreenMessageColour,
 
        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$IncludeSysError
 
 )
 
    process {
 
        # Stores Variables
        $LogsFolder      		     = 'Logs'
 

        # Date
        $DateNow = Get-Date -f g
        
        # Error Message
        $SysErrorMessage = $error[0].Exception.message
 
 
        # Check of log files exist for this session
        If ($Global:TransactionLog -eq $null) {$Global:TransactionLog = ".\TransactionLog_$((get-date).ToString('yyyyMMdd_HHmm')).csv"}
 
        
        # Create Directory Structure
        if (! (Test-Path ".\$LogsFolder")) {new-item -path .\ -name ".\$LogsFolder" -type directory | out-null}
 
 
        $TransactionLogScreen = [pscustomobject][ordered]@{}
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Date"-Value $DateNow 
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Task" -Value $Task
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Result" -Value $Result
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Error" -Value $ErrorMessage
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "SystemError" -Value $SysErrorMessage
        
       
        # Output to screen
       
        if  ($Result -match "Information|Warning" -and $ShowScreenMessage -eq "$true"){
 
        Write-host $TransactionLogScreen.Date  -NoNewline -ForegroundColor GREEN
        Write-host " | " -NoNewline
        Write-Host $TransactionLogScreen.Task  -NoNewline
        Write-host " | " -NoNewline
        Write-host $TransactionLogScreen.Result -ForegroundColor $ScreenMessageColour 
        }
 
       if  ($Result -eq "Error" -and $ShowScreenMessage -eq "$true" -and $IncludeSysError -eq "$false"){
       Write-host $TransactionLogScreen.Date  -NoNewline -ForegroundColor GREEN
       Write-host " | " -NoNewline
       Write-Host $TransactionLogScreen.Task  -NoNewline
       Write-host " | " -NoNewline
       Write-host $TransactionLogScreen.Result -ForegroundColor $ScreenMessageColour -NoNewline 
       Write-host " | " -NoNewline
       Write-Host $ErrorMessage  -ForegroundColor $ScreenMessageColour
       }
 
       if  ($Result -eq "Error" -and $ShowScreenMessage -eq "$true" -and $IncludeSysError -eq "$true"){
       Write-host $TransactionLogScreen.Date  -NoNewline -ForegroundColor GREEN
       Write-host " | " -NoNewline
       Write-Host $TransactionLogScreen.Task  -NoNewline
       Write-host " | " -NoNewline
       Write-host $TransactionLogScreen.Result -ForegroundColor $ScreenMessageColour -NoNewline 
       Write-host " | " -NoNewline
       Write-Host $ErrorMessage -NoNewline -ForegroundColor $ScreenMessageColour
       if (!$SysErrorMessage -eq $null) {Write-Host " | " -NoNewline}
       Write-Host $SysErrorMessage -ForegroundColor $ScreenMessageColour
       Write-Host
       }
   
        # Build PScustomObject
        $TransactionLogFile = [pscustomobject][ordered]@{}
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Date"-Value "$datenow"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Task"-Value "$task"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Result"-Value "$result"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Error"-Value "$ErrorMessage"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "SystemError"-Value "$SysErrorMessage"
 
 
        $TransactionLogFile | Export-Csv -Path ".\$LogsFolder\$TransactionLog" -Append -NoTypeInformation
 
 
        # Clear Error Messages
        $error.clear()
    }   
 
 
 }

# FUNCTION - Record Failed Events
function RecordFailed () {

    $Group | Out-File -Append "$FailedGroups\FailedGroups.csv" -Encoding ascii

}


# Creating Directory Structure

    If(!(Test-Path -Path $ExportDirectory )){
            Write-Host "  Creating Directory: $ExportDirectory"
        New-Item -ItemType directory -Path $ExportDirectory | Out-Null}

    If(!(Test-Path -Path $FullGroupExportDirectory )){
            Write-Host "  Creating Directory: $FullGroupExportDirectory"
        New-Item -ItemType directory -Path $FullGroupExportDirectory | Out-Null}

    If(!(Test-Path -Path $FailedGroups )){
            Write-Host "  Creating Directory: $FailedGroups"
            New-Item -ItemType directory -Path $FailedGroups | Out-Null
            New-Item -Path . -Name ".\$FailedGroups\FailedGroups.csv" -ItemType "file" -Value 'GroupName' | Out-Null
        Add-Content -path ".\$FailedGroups\FailedGroups.csv" -value ""}



If ($CreatePlaceHolder.IsPresent) {


    Write-Host "--------------------------------------------------------------------------" -ForegroundColor DarkBlue
    Write-host "                                                                          "
    Write-Host "            Working on Distribution Group $Group - PlaceHolder            " -ForegroundColor YELLOW
    Write-host "                                                         Cloud Operation  "
    Write-Host "--------------------------------------------------------------------------" -ForegroundColor DarkBlue
    Write-Host
     

    If (((Get-DistributionGroup $Group -ErrorAction 'SilentlyContinue').IsValid) -eq $true) {

        
        
        Try {
            WriteTransactionsLogs -Task "Finding $Group in Office 365" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            $OldDG = Get-DistributionGroup $Group -ea stop
            $oldDGPermissions = Get-RecipientPermission -Identity $Group -EA STOP }
        Catch {WriteTransactionsLogs -Task "Failed Finding $Group in Office 365" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
            RecordFailed
            Exit
        }

        # check if the script is running in cloud else stop
        if ($OldDG.DistinguishedName -match "DC=PROD,DC=OUTLOOK,DC=COM"){}
        Else {WriteTransactionsLogs -Task "Script looks to be running against Onprem objects, QUITTING" -Result ERROR -ErrorMessage "Wrong Platform" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
            Exit
        }

        # Build Strings and Structure attribute data
        [System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object {$Group = $Group.Replace($_,'_')}
        $OldName = [string]$OldDG.Name
        $OldDisplayName = [string]$OldDG.DisplayName
        $OldPrimarySmtpAddress = [string]$OldDG.PrimarySmtpAddress
        $OldAlias = [string]$OldDG.Alias
        $OldMembers = (Get-DistributionGroupMember $OldDG.Name -resultsize unlimited).DistinguishedName
        
        # Export data to directory for later use
        "EmailAddress" > "$ExportDirectory\$Group.csv"
        $OldDG.EmailAddresses >> "$ExportDirectory\$Group.csv"
        "x500:"+$OldDG.LegacyExchangeDN >> "$ExportDirectory\$Group.csv"
       
      
        Try {
            WriteTransactionsLogs -Task "Creating Group: Cloud-$OldDisplayName" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false

            New-DistributionGroup `
            -Name "Cloud-$OldName" `
            -Alias "Cloud-$OldAlias" `
            -DisplayName "Cloud-$OldDisplayName" `
	        -ManagedBy $OldDG.ManagedBy `
            -Members $OldMembers `
            -PrimarySmtpAddress "Cloud-$OldPrimarySmtpAddress" -ea stop  | Out-Null
        }
        Catch {WriteTransactionsLogs -Task "Failed Creating Group: Cloud-$OldDisplayName" -Result ERROR -ErrorMessage "Failed to Create Group" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
            Exit
        }

        Start-Sleep -Seconds 15


        Try {
            WriteTransactionsLogs -Task "Setting Values For: Cloud-$OldDisplayName" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
            Set-DistributionGroup `
            -Identity "Cloud-$OldName" `
            -AcceptMessagesOnlyFromSendersOrMembers $OldDG.AcceptMessagesOnlyFromSendersOrMembers `
            -RejectMessagesFromSendersOrMembers $OldDG.RejectMessagesFromSendersOrMembers -EA STOP `

            Set-DistributionGroup `
            -Identity "Cloud-$OldName" `
            -AcceptMessagesOnlyFrom $OldDG.AcceptMessagesOnlyFrom `
            -AcceptMessagesOnlyFromDLMembers $OldDG.AcceptMessagesOnlyFromDLMembers `
            -BypassModerationFromSendersOrMembers $OldDG.BypassModerationFromSendersOrMembers `
            -BypassNestedModerationEnabled $OldDG.BypassNestedModerationEnabled `
            -CustomAttribute1 $OldDG.CustomAttribute1 `
            -CustomAttribute2 $OldDG.CustomAttribute2 `
            -CustomAttribute3 $OldDG.CustomAttribute3 `
            -CustomAttribute4 $OldDG.CustomAttribute4 `
            -CustomAttribute5 $OldDG.CustomAttribute5 `
            -CustomAttribute6 $OldDG.CustomAttribute6 `
            -CustomAttribute7 $OldDG.CustomAttribute7 `
            -CustomAttribute8 $OldDG.CustomAttribute8 `
            -CustomAttribute9 $OldDG.CustomAttribute9 `
            -CustomAttribute10 $OldDG.CustomAttribute10 `
            -CustomAttribute11 $OldDG.CustomAttribute11 `
            -CustomAttribute12 $OldDG.CustomAttribute12 `
            -CustomAttribute13 $OldDG.CustomAttribute13 `
            -CustomAttribute14 $OldDG.CustomAttribute14 `
            -CustomAttribute15 $OldDG.CustomAttribute15 `
            -ExtensionCustomAttribute1 $OldDG.ExtensionCustomAttribute1 `
            -ExtensionCustomAttribute2 $OldDG.ExtensionCustomAttribute2 `
            -ExtensionCustomAttribute3 $OldDG.ExtensionCustomAttribute3 `
            -ExtensionCustomAttribute4 $OldDG.ExtensionCustomAttribute4 `
            -ExtensionCustomAttribute5 $OldDG.ExtensionCustomAttribute5 `
            -GrantSendOnBehalfTo $OldDG.GrantSendOnBehalfTo `
            -HiddenFromAddressListsEnabled $True `
            -MailTip $OldDG.MailTip `
            -MailTipTranslations $OldDG.MailTipTranslations `
            -MemberDepartRestriction $OldDG.MemberDepartRestriction `
            -MemberJoinRestriction $OldDG.MemberJoinRestriction `
            -ModeratedBy $OldDG.ModeratedBy `
            -ModerationEnabled $OldDG.ModerationEnabled `
            -RejectMessagesFrom $OldDG.RejectMessagesFrom `
            -RejectMessagesFromDLMembers $OldDG.RejectMessagesFromDLMembers `
            -ReportToManagerEnabled $OldDG.ReportToManagerEnabled `
            -ReportToOriginatorEnabled $OldDG.ReportToOriginatorEnabled `
            -RequireSenderAuthenticationEnabled $OldDG.RequireSenderAuthenticationEnabled `
            -SendModerationNotifications $OldDG.SendModerationNotifications `
            -SendOofMessageToOriginatorEnabled $OldDG.SendOofMessageToOriginatorEnabled `
            -BypassSecurityGroupManagerCheck -EA Stop
        }
        Catch {WriteTransactionsLogs -Task "Failed Setting Values For: Cloud-$OldDisplayName" -Result Information -ErrorMessage "Failed to Create Group" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        }

        # Setting any Recipient Permissions
        WriteTransactionsLogs -Task "Setting SendAs Permissions For: Cloud-$OldDisplayName" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        Foreach ($mailbox in $oldDGPermissions){

        Try {Add-RecipientPermission -identity "Cloud-$OldName" -AccessRights "Sendas" -trustee $mailbox.Trustee -Confirm:$false -EA stop -WarningAction SilentlyContinue | Out-Null }
        Catch {WriteTransactionsLogs -Task "Failed Adding SendAs Permissions For: Cloud-$OldDisplayName Trustee:$mailbox.Trustee " -Result Warning -ErrorMessage AddPermissionsError -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError true}
        }
    }   # Close of Group PlaceHolder          
               
    Else {
        Write-Host "  ERROR: The distribution group '$Group' was not found" -ForegroundColor Red
        RecordFailed
        Write-Host
    }
}
ElseIf ($Finalize.IsPresent) {



    Write-Host "--------------------------------------------------------------------------" -ForegroundColor DarkBlue
    Write-host ""
    Write-Host "            Working on Distribution Group $Group - Finalize               " -ForegroundColor YELLOW
    Write-host "                                                         Cloud Operation  "
    Write-Host "--------------------------------------------------------------------------" -ForegroundColor DarkBlue
    Write-Host


      
    Try {
        WriteTransactionsLogs -Task "Collecting Group information for $Group" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $TempDG = Get-DistributionGroup "Cloud-$Group" -EA Stop
        $TempPrimarySmtpAddress = $TempDG.PrimarySmtpAddress
    }
    Catch {WriteTransactionsLogs -Task "Failed Collecting Group information for $Group" -Result Information -ErrorMessage "Find group ERROR in 365" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        Exit
    } 

        
    # check if the script is running in cloud else stop
    if ($TempDG.DistinguishedName -match "DC=PROD,DC=OUTLOOK,DC=COM"){}
    Else {WriteTransactionsLogs -Task "Script looks to be running against Onprem objects, QUITTING" -Result ERROR -ErrorMessage "Wrong Platform" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        exit
    }


    [System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object {$Group = $Group.Replace($_,'_')}

    $OldAddresses = @(Import-Csv "$ExportDirectory\$Group.csv")
    $NewAddresses = $OldAddresses | ForEach-Object {$_.EmailAddress.Replace("X500","x500")}
    $NewDGName = $TempDG.Name.Replace("Cloud-","")
    $NewDGDisplayName = $TempDG.DisplayName.Replace("Cloud-","")
    $NewDGAlias = $TempDG.Alias.Replace("Cloud-","")
    $NewPrimarySmtpAddress = ($NewAddresses | Where-Object {$_ -clike "SMTP:*"}).Replace("SMTP:","")

    Try {
        WriteTransactionsLogs -Task "Updating $Group with recorded values" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        Set-DistributionGroup `
        -Identity $TempDG.Name `
        -Name $NewDGName `
        -Alias $NewDGAlias `
        -DisplayName $NewDGDisplayName `
        -PrimarySmtpAddress $NewPrimarySmtpAddress `
        -HiddenFromAddressListsEnabled $False `
        -BypassSecurityGroupManagerCheck -EA Stop

        Set-DistributionGroup `
        -Identity $NewDGName `
        -EmailAddresses @{Add=$NewAddresses} `
        -BypassSecurityGroupManagerCheck -EA Stop

        Set-DistributionGroup `
        -Identity $NewDGName `
        -EmailAddresses @{Remove=$TempPrimarySmtpAddress} `
        -BypassSecurityGroupManagerCheck -EA STOP
    }
    Catch {WriteTransactionsLogs -Task "Failed Updating $Group with recorded values" -Result ERROR -ErrorMessage UpdateError -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
        exit}
    } # Close of Group Finalize

ElseIf ($Contact.IsPresent) {

    

    If (((Get-DistributionGroup $Group -ErrorAction 'SilentlyContinue').IsValid) -eq $true) {



    Write-Host "--------------------------------------------------------------------------" -ForegroundColor DarkBlue
    Write-host ""
    Write-Host "            Working on Distribution Group $Group - Contact Create         " -ForegroundColor YELLOW
    Write-host "                                                        Onprem Operation  "
    Write-Host "--------------------------------------------------------------------------" -ForegroundColor DarkBlue
    Write-Host
     

    # Collect group info
    Try {
        WriteTransactionsLogs -Task "Found Distribution $Group" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $GroupData = Get-DistributionGroup $group -EA STOP

        #Export Data
        $GroupData | Export-Clixml "$FullGroupExportDirectory\$group.xml"

        #Create simple strings
        [string]$CurrentSMTP        = $groupdata.PrimarySmtpAddress
        [string]$CurrentName        = $groupdata.Name
        [string]$CurrentDisplayName = $groupdata.DisplayName
        [string]$CurrentAlias       = $groupdata.Alias
        $GroupEmailAddresses        = $groupdata.EmailAddresses
        WriteTransactionsLogs -Task "Collected Distribution Group details for $Group" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}
    Catch {WriteTransactionsLogs -Task "Failed Collecting Distribution Group details for $Group" -Result ERROR -ErrorMessage "Failed Getting DL Details" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
        exit
    }


    # check if the script is running on premise else stop
    if ($Groupdata.DistinguishedName -Notmatch "DC=PROD,DC=OUTLOOK,DC=COM"){}
    Else {WriteTransactionsLogs -Task "Script looks to be running against cloud objects, QUITTING" -Result ERROR -ErrorMessage "Wrong Platform" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        exit
    }


    # Create new object data to be stamped
    [string]$NewSMTP        = "old_$CurrentSMTP"
    [string]$NewName        = "old_$CurrentName"
    [string]$NewDisplayName = "old_$CurrentDisplayName"
    [string]$NewAlias       = "old_$CurrentAlias"

    # Find Onmicrosorft address to be used as target 
    $TargetOnMicrosoft = ($GroupEmailAddresses | Where-Object {$_ -clike "smtp:*mail.onmicrosoft.com"}) | Select-Object -ExpandProperty smtpAddress
    If ($TargetOnMicrosoft){WriteTransactionsLogs -Task "Found Onmicrosoft Address for $Group : $TargetOnMicrosoft" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}
    If (!($TargetOnMicrosoft)){WriteTransactionsLogs -Task "No Onmicrosoft Address for $Group - Script will exist" -Result ERROR -ErrorMessage "No Onmicrosoft Address Found" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        Exit
    }
            

    try {
        Set-DistributionGroup `
        -Identity $groupData.DistinguishedName`
        -Alias $NewAlias `
        -DisplayName $NewDisplayName `
        -PrimarySmtpAddress $NewSMTP `
        -HiddenFromAddressListsEnabled $False `
        -EmailAddressPolicyEnabled $False `
        -CustomAttribute5  'Do-not-sync'`
        -BypassSecurityGroupManagerCheck `
        -DomainController $DCServer -EA STOP
        WriteTransactionsLogs -Task "Set DistributionGroup details to Old / Not Synced" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    }
    Catch {WriteTransactionsLogs -Task "Failed to Set DistributionGroup details to Old / Not Synced" -Result ERROR -ErrorMessage "Failed to Set DL details to old" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
        Exit
    }

    try {
        Set-DistributionGroup `
        -Name $NewName `
        -Identity $groupData.DistinguishedName `
        -EmailAddresses @{Remove=$TargetOnMicrosoft,$CurrentSMTP} `
        -BypassSecurityGroupManagerCheck `
        -DomainController $DCServer -EA STOP
        WriteTransactionsLogs -Task "Removed Addresses: $TargetOnMicrosoft,$CurrentSMTP" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    }
    Catch {WriteTransactionsLogs -Task "Failed Removing Addresses: $TargetOnMicrosoft,$CurrentSMTP from $Group" -Result ERROR -ErrorMessage "Failed removing Addresses" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true}

    # Sleep for a little
    WriteTransactionsLogs -Task "Sleeping for a little as AD is old" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
    Start-Sleep -s 5

    # Create Contact with forwarder to cloud group 
    Try {
        writeTransactionsLogs -Task "Creating New contact based on Group details" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $NewContact = New-MailContact -Name  $CurrentName -OrganizationalUnit $ContactGroupOU -DisplayName  $CurrentDisplayname -PrimarySmtpAddress $CurrentSMTP -ExternalEmailAddress $TargetOnMicrosoft -Alias $CurrentAlias -DomainController $DCServer -EA STOP
        }
    Catch {writeTransactionsLogs -Task "Failed to create mail contact object" -Result Information -ErrorMessage "Failed creating contact" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false
        Exit
    }

    Try {
        WriteTransactionsLogs -Task "Setting Additional Mail Contact object information" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        Start-Sleep 1
        Set-MailContact -identity $NewContact.DistinguishedName -HiddenFromAddressListsEnabled:$true -CustomAttribute5 'DO-NOT-SYNC' -DomainController $DCServer -EA STOP
        Set-Contact -identity $NewContact.DistinguishedName -Notes "This contact is used to support the $CurrentName group located in 365. " -DomainController $DCServer -EA STOP
        Set-ADObject -Identity $NewContact.DistinguishedName -add @{Notes="This contact is used to support the $CurrentName group located in 365."; Description="This contact is used to support the $CurrentName group located in 365."} -Server $DCServer -EA STOP
        writeTransactionsLogs -Task "Completed Mail Contact object for $Group" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false}
        Catch {writeTransactionsLogs -Task "Failed Setting additional Mail Contact object Information" -Result ERROR -ErrorMessage "Failed updating object" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
    }

    }
    Else {
        writeTransactionsLogs -Task "Distribution Group $Group - NOT FOUND" -Result Information -ErrorMessage "NOT Found Group In AD" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError false
        RecordFailed
    }

    } # Close of Contact option
    
Else {
    Write-Host "  ERROR: No options selected, please use '-CreatePlaceHolder' or '-Finalize' or '-Contact'" -ForegroundColor Red
    Write-Host
}