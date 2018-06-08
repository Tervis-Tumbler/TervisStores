$ModulePath = (Get-Module -ListAvailable TervisStores).ModuleBase
. $ModulePath\Definition.ps1

function Get-iVMSGitRepositoryPath {
    $ADDomain = Get-ADDomain -Current LocalComputer
    "\\$($ADDomain.DNSRoot)\applications\GitRepository\iVMS-4200"
}

function Install-TervisiVMSConfiguration {
    param (
        [Parameter(Mandatory)]$ComputerName
    )
    begin {
        $iVMSConfigurationFilePath = Join-Path -Path (Get-iVMSGitRepositoryPath) -ChildPath "iVMS-4200Configuration.zip"
        $LocaliVMSInstallPath = "C:\Program Files\iVMS-4200 Station\iVMS-4200\iVMS-4200 Client"
    }
    process {
        $RemoteiVMSInstallPath = $LocaliVMSInstallPath | ConvertTo-RemotePath -ComputerName $ComputerName
        If (Test-Path -Path $RemoteiVMSInstallPath) {
            Expand-Archive -Path $iVMSConfigurationFilePath -DestinationPath $RemoteiVMSInstallPath -Force 
        } else {
            Write-Warning "iVMS-4200 not installed on $ComputerName"
        }
    }
}

function Get-StoreNameFromADUser {
    $OrganizationalUnit = Get-ADOrganizationalUnit -Filter * | 
    Where-Object DistinguishedName -Match "OU=Back Office,OU=Remote,OU=Users"

    Get-ADUser -SearchBase $OrganizationalUnit.DistinguishedName -Filter * |
    Select-Object -ExpandProperty GivenName
}

function Get-StoreEmailLocalPartFromName {
    $StoreNames = Get-StoreNameFromADUser
    $StoreNamesWithoutSpaces = $StoreNames.replace(" ", "")
    
    $StoreNamesWithoutSpaces
}

function Get-StoreEmailAddressesFromName {
    Get-StoreEmailLocalPartFromName | 
    ForEach-Object { 
        "$_@TervisStore.com"
    }
}

function New-MigaduStoreEmailBox {
    param (
        [Parameter(Mandatory)]$XAuthorizationToken,
        [Parameter(Mandatory)]$XAuthorizationEmail
    )
    foreach ($Store in $StoreDefinition) {
        $Credential = Get-PasswordstateCredential -PasswordID $Store.EmailAccountPasswordStateID
        $Domain = $Credential.UserName -split "@" | select -First 1 -Skip 1
        $EmailAddressLocalPart = $Credential.UserName -split "@" | select -First 1
        $MigaduMailbox = Get-MigaduMailbox -Domain $Domain -EmailAddressLocalPart $EmailAddressLocalPart -XAuthorizationToken $XAuthorizationToken -XAuthorizationEmail $XAuthorizationEmail -ErrorAction SilentlyContinue

        if (-not $MigaduMailbox) {
            New-MigaduMailbox -XAuthorizationToken $XAuthorizationToken -XAuthorizationEmail $XAuthorizationEmail -Domain $Domain -EmailAddressLocalPart $EmailAddressLocalPart -DisplayName "$($Store.Name) Store" -Password $Credential.GetNetworkCredential().password
        }
    }
}

function Install-Office2016OnBackOfficeComputers {
    $BackOfficeComputers = Get-BackOfficeComputers -Online

    Start-ParallelWork -ScriptBlock {
        param($Parameter)
        Install-TervisOffice2016VLPush -ComputerName $Parameter
    } -Parameters $BackOfficeComputers
}

function Set-StoreMigaduMailboxEnvironmentVariablesOnAllBackOffice {
    process {
        $BackOfficeComputerNames = Get-BackOfficeComputers -Online

        Foreach ($ComputerName in $BackOfficeComputerNames) {
            Set-StoreMigaduMailboxEnvironmentVariables -ComputerName $ComputerName
        }
    }
}

function Set-StoreMigaduMailboxEnvironmentVariables {
    param (
        [Parameter(ValueFromPipelineByPropertyName)]$ComputerName
    )
    process {
        $StoreNumber = $ComputerName.substring(0,4)
        $StoreDefinition = Get-TervisStoreDefinition -Number $StoreNumber
        Set-EnvironmentVariable -Name MigaduMailboxDisplayName -ComputerName $ComputerName -Value "$($StoreDefinition.Name) Store" -Target Machine
        Set-EnvironmentVariable -Name MigaduEmailAddress -ComputerName $ComputerName -Value $StoreDefinition.EmailAddress -Target Machine
    }
}

function Get-TervisStoreDefinition {
    param (
        $Name,
        $Number
    )
    $StoreDefinition |
    Where-Object { -Not $Name -or $_.Name -Match $Name } |
    Where-Object { -Not $Number -or $_.Number -Match $Number } |
    Add-TervisStoreDefinitionCustomProperty -PassThru
}

function Add-TervisStoreDefinitionCustomProperty {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$Object,
        [Switch]$PassThru
    )
    process {
        $Object |
        Add-Member -MemberType ScriptProperty -Name BackOfficeUserCredential -Force -Value {
            Get-PasswordstateCredential -PasswordID $This.BackOfficeAccountPasswordStateID
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name BackOfficeUserName -Force -Value {
            $This.BackOfficeUserCredential.UserName
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name BackOfficeADUser -Force -Value {
            Get-AdUser -Identity $This.BackOfficeUserName -Properties *
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name MigaduMailboxCredential -Force -Value {
            Get-PasswordstateCredential -PasswordID $This.EmailAccountPasswordStateID
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name EmailAddress -Force -Value {
            $This.MigaduMailboxCredential |
            Select-Object -ExpandProperty UserName
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name MailContact -Force -Value {
            Import-TervisExchangePSSession
            Get-ExchangeMailContact -Identity $This.EmailAddress
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name TervisDotComDistributionGroup -Force -Value {
            Import-TervisExchangePSSession
            Get-ExchangeDistributionGroup -Identity "$($This.Name) tervis.com email address Distribution Group"
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name ExchangeMailbox -Force -Value {
            Import-TervisExchangePSSession
            Get-ExchangeMailbox -Identity $This.BackOfficeUserName
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name MailContactADObject -Force -Value {
            $EmailAddress = $This.EmailAddress
            Get-ADObject -Filter { Mail -eq $EmailAddress } -Properties *
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name Computer -Force -Value {
            $StoreNumber = $This.Number
            $FilterValue = "$StoreNumber*"
            Get-ADComputer -Filter {Name -like $FilterValue}
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name BackOffice -Force -Value {
            $This.Computer |
            Where-Object Name -Match "BO"
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name Register -Force -Value {
            $This.Computer |
            Where-Object Name -Match "POS"
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name NumberOfBackOffice -Force -Value {
            $This.BackOffice | Measure-Object | Select-Object -ExpandProperty Count
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name NumberOfRegister -Force -Value {
            $This.Register | Measure-Object | Select-Object -ExpandProperty Count
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name OldSMTPAddressWithOldInName -Force -Value {
            "$($This.BackOfficeADUser.SamAccountName)_Old@tervis.com"
        }

        $Object |
        Where-Object {-not $_.OldSMTPAddress} |
        Add-Member -MemberType ScriptProperty -Name OldSMTPAddress -Force -Value {
            "$($This.BackOfficeADUser.SamAccountName)@tervis.com"
        }

        if ($PassThru) { $Object }
    }
}

function Install-OutlookMigaduMailProfile {
& "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE" /importprf "\\tervis.prv\departments\Stores\Stores Shared\Migadu\OutlookMigaduProfile.PRF" 
& "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE" /importprf "\\tervis.prv\departments\Stores\Stores Shared\Migadu\OutlookMigaduProfile.PRF"
& "C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE" /importprf "\\tervis.prv\departments\Stores\Stores Shared\Migadu\OutlookMigaduProfile.PRF"
}

function Edit-OutlookMigaduMailProfile {
& "\\tervis.prv\applications\Installers\Microsoft\Office 2016 Professional Plus Volume Licensing Edition\Office Professional Plus 2016 32Bit Volume Licensing Edition\setup.exe" /admin
}

function Invoke-InstallEnvironmentVariablesAndTriggerRebootForTomorrow {
    param (
        [Parameter(Mandatory)]$ComputerName
    )
    Set-StoreMigaduMailboxEnvironmentVariables -ComputerName $ComputerName
    $TomorrowAt430am = [DateTime]::Today.AddDays(1).AddHours(4).AddMinutes(30)
    Restart-TervisComputer -DateTimeOfRestart $TomorrowAt430am -ComputerName $ComputerName -Message "Rebooted needed to complete change made by IT"
}

function Restart-BackOfficeComputersNotRebootedSinceDate {
    param (
        [Parameter(Mandatory)][DateTime]$DateTimeOfRestart,
        [DateTime]$HaventRebootedSinceDate,
        $Message
    )
    $BackOfficeComputerNames = Get-BackOfficeComputers -Online
    $BackOfficeComputerNames | ForEach-Object {
        Restart-TervisComputerIfNotRebootedSinceDateTime -DateTimeOfRestart $DateTimeOfRestart -HaventRebootedSinceDate $HaventRebootedSinceDate -ComputerName $_ -Message $Message
    }
}

function New-StoreEmailAddressContact {
    $StoreDefinitions = Get-TervisStoreDefinition
    Import-TervisExchangePSSession
    
    foreach ($StoreDefinition in $StoreDefinitions) {
        $OrganizationalUnit = Get-ADOrganizationalUnit -filter * | where DistinguishedName -Match "OU=Contacts,OU=Stores"
        $DisplayName = "$($StoreDefinition.Name) Store"
        New-ExchangeMailContact -FirstName $StoreDefinition.Name -LastName Store -Name $DisplayName -ExternalEmailAddress $StoreDefinition.EmailAddress -OrganizationalUnit $OrganizationalUnit.Name
    }
}

function Sync-StoreDistributionGroupsWithContacts {   
    $StoreDefinitions = Get-TervisStoreDefinition
    $BackOfficeUserNames = $StoreDefinitions.BackOfficeUserName
    
    $Region1Stores = Get-ADGroupMember -Identity "Region 1 Stores" | 
    Where-Object SamAccountName -in $BackOfficeUserNames |
    Select-Object -ExpandProperty SamAccountName 
    
    $Region2Stores = Get-ADGroupMember -Identity "Region 2 Stores" | 
    Where-Object SamAccountName -in $BackOfficeUserNames |
    Select-Object -ExpandProperty SamAccountName

    
    $StoreDefinitions | 
    ForEach-Object {
        "$($_.Name) tervis.com email address Distribution Group"
    } |
    New-TervisDistributionGroup

    foreach ($StoreDefinition in $StoreDefinitions) {
        $StoreDefinition.TervisDotComDistributionGroup | 
        Add-ExchangeDistributionGroupMember -Member $StoreDefinition.MailContactADObject.DistinguishedName -ErrorAction SilentlyContinue
    }

}

function Move-StoreTervisDotComAddressesToDistributionGroup {
    $StoreDefinitions = Get-TervisStoreDefinition
    foreach ($StoreDefinition in $StoreDefinitions) {
        $TervisDotComEmailAddressSMTPForm = $StoreDefinition.ExchangeMailbox.EmailAddresses |
        Where-Object {$_ -match "tervis.com"}
        
        if ($TervisDotComEmailAddressSMTPForm) {
            $OnMicrosoftEmailAddress = (
                $StoreDefinition.ExchangeMailbox.EmailAddresses |
                Where-Object {$_ -match "mail.onmicrosoft.com"}
            ) -split ":" |
            Select-Object -First 1 -Skip 1
        
            $StoreDefinition.ExchangeMailbox | 
            Set-ExchangeMailbox -PrimarySmtpAddress $OnMicrosoftEmailAddress -EmailAddressPolicyEnabled:$false
        
            $StoreDefinition.ExchangeMailbox | 
            Set-ExchangeMailbox -EmailAddresses @{Remove = $TervisDotComEmailAddressSMTPForm} -EmailAddressPolicyEnabled:$false
             #The next line errors if something internal to exchange doesn't have enough time to realize the address is no longer on the mailbox

            $TervisDotComEmailAddress = $TervisDotComEmailAddressSMTPForm -split ":" |
            Select-Object -First 1 -Skip 1

            do {
                Start-Sleep -Seconds 10
                $StoreDefinition.TervisDotComDistributionGroup |
                Set-ExchangeDistributionGroup -PrimarySmtpAddress $TervisDotComEmailAddress -EmailAddressPolicyEnabled:$false
            } while ($StoreDefinition.TervisDotComDistributionGroup.PrimarySmtpAddress -ne $TervisDotComEmailAddress)
        }
    }
}

function Set-StoreDistributionGroupForTervisDotComAddressToHaveTervisDotComAddress {
    $StoreDefinitions = Get-TervisStoreDefinition
    foreach ($StoreDefinition in $StoreDefinitions) {
        $StoreDefinition.TervisDotComDistributionGroup |
        Set-ExchangeDistributionGroup -PrimarySmtpAddress $StoreDefinition.OldSMTPAddress -EmailAddressPolicyEnabled:$false
    }
}

function Set-StoreBackOfficeADUserUPN {
    $StoreDefinitions = Get-TervisStoreDefinition
    foreach ($StoreDefinition in $StoreDefinitions) {
        if ($StoreDefinition.BackOfficeADUser.UserPrincipalName -notmatch "@tervis0.onmicrosoft.com") {
            $StoreDefinition.BackOfficeADUser |
            Set-ADUser -UserPrincipalName "$($StoreDefinition.BackOfficeADUser.SamAccountName)@tervis0.onmicrosoft.com"
        }
    }
}

function Set-StoreTervisDotComAddressWithOldAsPrimaryForMailbox {
    $StoreDefinitions = Get-TervisStoreDefinition
    foreach ($StoreDefinition in $StoreDefinitions) {        
        if ($StoreDefinition.ExchangeMailbox.PrimarySmtpAddress -eq $StoreDefinition.OldSMTPAddress) {        
            $StoreDefinition.ExchangeMailbox | 
            Set-ExchangeMailbox -PrimarySmtpAddress $StoreDefinition.OldSMTPAddressWithOldInName -EmailAddressPolicyEnabled:$false
        }
    }
}

function Remove-StoreTervisDotComAddressFromMailbox {
    $StoreDefinitions = Get-TervisStoreDefinition
    foreach ($StoreDefinition in $StoreDefinitions) {
        $TervisDotComEmailAddressSMTPForm = $StoreDefinition.ExchangeMailbox.EmailAddresses |
        Where-Object {$_ -match $StoreDefinition.OldSMTPAddress}

        if ($TervisDotComEmailAddressSMTPForm) {
            $StoreDefinition.ExchangeMailbox | 
            Set-ExchangeMailbox -EmailAddresses @{Remove ="smtp:$($StoreDefinition.OldSMTPAddress)" } -EmailAddressPolicyEnabled:$false
        }
    }
}


function Remove-StoreTervisDotComAddressesFromDistributionGroup {
    $StoreDefinitions = Get-TervisStoreDefinition
    foreach ($StoreDefinition in $StoreDefinitions) {
        $TervisDotComEmailAddressSMTPForm = $StoreDefinition.TervisDotComDistributionGroup.EmailAddresses |
        Where-Object {$_ -match $StoreDefinition.OldSMTPAddress}
        
        if ($TervisDotComEmailAddressSMTPForm) {
            $OnMicrosoftEmailAddress = (
                $StoreDefinition.TervisDotComDistributionGroup.EmailAddresses |
                Where-Object {$_ -match "mail.onmicrosoft.com"}
            ) -split ":" |
            Select-Object -First 1 -Skip 1
        
            $StoreDefinition.TervisDotComDistributionGroup | 
            Set-ExchangeDistributionGroup -PrimarySmtpAddress $OnMicrosoftEmailAddress -EmailAddressPolicyEnabled:$false
        
            $StoreDefinition.TervisDotComDistributionGroup |
            Set-ExchangeDistributionGroup -EmailAddresses @{Remove = $TervisDotComEmailAddressSMTPForm} -EmailAddressPolicyEnabled:$false
        }
    }
}

function Add-OldTervisDotComEmailAddressesToOldMailbox {
    $StoreDefinitions = Get-TervisStoreDefinition

    foreach ($StoreDefinition in $StoreDefinitions) {
        $StoreDefinition.ExchangeMailbox | 
        Set-ExchangeMailbox -PrimarySmtpAddress $StoreDefinition.OldSMTPAddress -EmailAddressPolicyEnabled:$false
    }
}

function Update-GroupsContainingStoreTervisDotComAddressToUseMailContact {
    $StoreDefinitions = Get-TervisStoreDefinition
    foreach ($StoreDefinition in $StoreDefinitions) {
        $MailEnabledGroups =  $StoreDefinition.BackOfficeADUser.MemberOf | 
        Get-ADGroup -Properties Mail |
        Where {$_.Mail}

        foreach ($Group in $MailEnabledGroups) {
            $Group | Set-ADGroup -Add @{Member = $StoreDefinition.MailContactADObject.DistinguishedName}
            $Group | Remove-ADGroupMember -Members $StoreDefinition.BackOfficeADUser -Confirm:$false
        }
    }
}

function Set-StoresOldMailboxToHiddinInGAL {
    $StoreDefinitions = Get-TervisStoreDefinition
    $StoreDefinitions.ExchangeMailbox | set-ExchangeMailbox -HiddenFromAddressListsEnabled $true
}

function Set-StoresADAccountsAsMembersOfTervisEveryone {
    $StoreDefinitions = Get-TervisStoreDefinition
    Get-ADGroup -Identity "Tervis - Everyone" |
    Add-ADGroupMember -Members $StoreDefinitions.BackOfficeADUser 

    Get-ADGroup -Identity "Stores" |
    Add-ADGroupMember -Members $StoreDefinitions.BackOfficeADUser 

    $Members = Get-ADGroup -Identity "Stores" | 
    Get-ADGroupMember 
    Add-ADGroupMember -Identity "Stores-1367133941" -Members $Members
    
    Get-ADGroup -Identity "Stores-1367133941" | Get-ADGroupMember

    Disable-ExchangeDistributionGroup -Identity "Stores"
}

function Invoke-RemoveStoresFromDistroGroupAndAddMailContact {
    $StoreDefinitions = Get-TervisStoreDefinition
    $Group = Get-ADGroup -Identity "Stores-1367133941"
    
    foreach ($StoreDefinition in $StoreDefinitions) {        
        $Group | Set-ADGroup -Add @{Member = $StoreDefinition.MailContactADObject.DistinguishedName}
        $Group | Remove-ADGroupMember -Members $StoreDefinition.BackOfficeADUser -Confirm:$false
    }
}

function Add-PrimaryEmailAddressToOldStoreMailboxes {
    $StoreDefinitions = Get-TervisStoreDefinition
    foreach ($StoreDefinition in $StoreDefinitions) {        
        $StoreDefinition.ExchangeMailbox |
        Set-ExchangeMailbox -PrimarySmtpAddress "$($StoreDefinition.BackOfficeADUser.SamAccountName)_Old@Tervis.com" -EmailAddressPolicyEnabled:$false
    }


    $StoreDefinition.TervisDotComDistributionGroup
}

function Invoke-GivexDeployment {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    process {
        Install-TervisChocolatey -ComputerName $ComputerName
        Install-GivexRMSPlugin -ComputerName $ComputerName
        Add-GivexRMSTenderType -ComputerName $ComputerName
        #Remove-StandardGiftCardTenderType
        #Add-GivexBalanceCustomPOSButton
        #Add-GivexAdminCustomPOSButton
        #Install-GivexReceipt
        #Install-GivexGcmIniFile
    }
}

function Install-GivexRMSPlugin {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    begin {
        $PackageSource = "\\$env:USERDNSDOMAIN\applications\Chocolatey\givexrmsplugin.1.4.0.261702.nupkg"
        $DestinationLocal = "C:\ProgramData\Tervis\ChocolateyPackage\givexrmsplugin.1.4.0.261702.nupkg"
    }
    process {
        Copy-ItemToRemoteComputerWithBitsTransfer -ComputerName $ComputerName -Source $PackageSource -DestinationLocal $DestinationLocal
        
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            choco install chocolatey-uninstall.extension -y
            choco install $Using:DestinationLocal -y
        }
    }
}

function Add-GivexRMSTenderType {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    process {
        $GivexTenderTypeParameters = @{
            ComputerName = $ComputerName
            Description = "Givex Gift Certificate"
            Code = "GIVEX"
            DoNotPopCashDrawer = 1
            AllowMultipleEntries = 1
            DisplayOrder = 20
        }
    
        Add-TervisRMSTenderType @GivexTenderTypeParameters
    }
}
