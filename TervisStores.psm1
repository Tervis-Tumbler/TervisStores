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
        $Credential = Get-PasswordstatePassword -AsCredential -ID $Store.EmailAccountPasswordStateID
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
            Get-PasswordstatePassword -AsCredential -ID $This.BackOfficeAccountPasswordStateID
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name BackOfficeUserName -Force -Value {
            $This.BackOfficeUserCredential.UserName
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name BackOfficeADUser -Force -Value {
            Get-AdUser -Identity $This.BackOfficeUserName -Properties *
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name MigaduMailboxCredential -Force -Value {
            Get-PasswordstatePassword -AsCredential -ID $This.EmailAccountPasswordStateID
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

function Invoke-GivexDeploymentToBackOfficeComputer {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]$ComputerObject
    )
    
    Write-Verbose "Getting database names"
    $ComputerObject | ForEach-Object {
        $DatabaseName = Get-RMSDatabaseName -ComputerName $_.ComputerName | Select-Object -ExpandProperty RMSDatabaseName
        $_ | Add-Member -MemberType NoteProperty -Name DatabaseName -Value $DatabaseName -Force
    }
    $ComputerObject | Add-GivexRMSTenderType -Verbose
    $ComputerObject | Remove-StandardGiftCardTenderType -Verbose
    $ComputerObject | Add-GivexBalanceCustomPOSButton -Verbose
    $ComputerObject | Add-GivexAdminCustomPOSButton -Verbose
    $ComputerObject | Set-GivexRMSItemProperties -Verbose
}

function Invoke-GivexDeploymentToRegisterComputer {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]$ComputerObject,
        [switch]$DeltaEnvironment
    )
 
    $ComputerObject | Install-TervisChocolatey
    $ComputerObject | Install-GivexRMSPlugin
    if ($DeltaEnvironment) {
        $ComputerObject | Install-GivexGcmIniFile_DEV
    } else {
        $ComputerObject | Install-GivexGcmIniFile
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
        Write-Verbose "$ComputerName - Installing Givex driver"
        Copy-ItemToRemoteComputerWithBitsTransfer -ComputerName $ComputerName -Source $PackageSource -DestinationLocal $DestinationLocal        
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            choco install chocolatey-uninstall.extension -y
            choco install $Using:DestinationLocal -y
        }
    }
}

function Add-GivexRMSTenderType {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$DatabaseName
    )
    process {
        Write-Verbose "$ComputerName - Adding RMS Tender Type for Givex"
        $GivexTenderTypeParameters = @{
            ComputerName = $ComputerName
            DatabaseName = $DatabaseName
            Description = "Givex Gift Certificate"
            Code = "GIVEX"
            DoNotPopCashDrawer = 1
            AllowMultipleEntries = 1
            DisplayOrder = 20
        }
        
        Add-TervisRMSTenderType @GivexTenderTypeParameters
    }
}

function Remove-StandardGiftCardTenderType {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$DatabaseName
    )
    begin {
        $Query = "DELETE FROM Tender WHERE Description = 'Gift Card'"
    }
    process {
        Write-Verbose "$ComputerName - Removing old gift card Tender Type"        
        Invoke-RMSSQL -DataBaseName $DatabaseName -SQLServerName $ComputerName -Query $Query
    }
}

function Add-GivexBalanceCustomPOSButton {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$DatabaseName
    )
    process {
        Write-Verbose "$ComputerName - Adding Givex Balance button"
        $GivexBalanceCustomPOSButtonParameters = @{
            ComputerName = $ComputerName
            DatabaseName = $DatabaseName
            Number = 4
            Style = 7
            Command = "RmsGcm.RmsBalance"
            Description = "Givex Balance Button"
            Picture = "0x47494638396120013200F70000FFFFFFFFFFCCFFFF99FFFF66FFFF33FFFF00FFCCFFFFCCCCFFCC99FFCC66FFCC33FFCC00FF99FFFF99CCFF9999FF9966FF9933FF9900FF66FFFF66CCFF6699FF6666FF6633FF6600FF33FFFF33CCFF3399FF3366FF3333FF3300FF00FFFF00CCFF0099FF0066FF0033FF0000CCFFFFCCFFCCCCFF99CCFF66CCFF33CCFF00CCCCFFCCCCCCCCCC99CCCC66CCCC33CCCC00CC99FFCC99CCCC9999CC9966CC9933CC9900CC66FFCC66CCCC6699CC6666CC6633CC6600CC33FFCC33CCCC3399CC3366CC3333CC3300CC00FFCC00CCCC0099CC0066CC0033CC000099FFFF99FFCC99FF9999FF6699FF3399FF0099CCFF99CCCC99CC9999CC6699CC3399CC009999FF9999CC9999999999669999339999009966FF9966CC9966999966669966339966009933FF9933CC9933999933669933339933009900FF9900CC99009999006699003399000066FFFF66FFCC66FF9966FF6666FF3366FF0066CCFF66CCCC66CC9966CC6666CC3366CC006699FF6699CC6699996699666699336699006666FF6666CC6666996666666666336666006633FF6633CC6633996633666633336633006600FF6600CC66009966006666003366000033FFFF33FFCC33FF9933FF6633FF3333FF0033CCFF33CCCC33CC9933CC6633CC3333CC003399FF3399CC3399993399663399333399003366FF3366CC3366993366663366333366003333FF3333CC3333993333663333333333003300FF3300CC33009933006633003333000000FFFF00FFCC00FF9900FF6600FF3300FF0000CCFF00CCCC00CC9900CC6600CC3300CC000099FF0099CC0099990099660099330099000066FF0066CC0066990066660066330066000033FF0033CC0033990033660033330033000000FF0000CC000099000066000033000000C0C2B8C0C3B0C3D93FC3D747C3D64FC2D457C2D35FC2D167C2D070C2CE78C2CD7FC1CB87C1C98FC1C898C1C6A0C1C5A85F5C5C5855556C69696562627F7D7D7977777270708C8A8A868484A6A5A59F9E9E999898929191C0C0C0B9B9B9B3B3B3ACACACFFFFFF00000000000000000000000000000000000021F904010000F9002C00000000200132000008FF00EB091C48B0A0C18308132A5CC8B0A1C38710234A9C48B1A2C58B18336ADCC8B1A3C78F20438A1C49B2A4C9932853AA5CC9B2A5CB973063CA9C49B3A6CD9B3873EADCC9B3A7CF9F40830A1D5A539E3B77EFEC115DCAB429C179E9A2A673E7B4AAD59FEBA446C57775E8BD7760C1DEEB8A536BD47864211E5DBB16DE3BB425E399859B966656AD5CEB3634CB175DD29172B5D2D51B539E567684F7F2E5BB2E2FC8C05207277EF9F5DDBCC90C172F46A7F4F15CCC27E7C9036D51F362C49E0593DE682FDE3CCBF1C6D6839AEE324B7CAFC3CA938D709C3871E722623347FC5C706CC1336B5D77F4AE59DE1D219F5D8D719EBAC5EAE8A1930A4FA074AD02DB99FFED9E109ED97705E73937ABCE764172DAE26BF3960DE139F9F1BB6513C7EDF7B67AE378A39854E879B79D56EE45F7197514D9E38E698B75F79D540219B69C42EBA5C31B3ED7413855670379839F36DBD4575036DBE0D70D360209688E36E57423A0720412F45D810ADD234F589225F45D8F04B5A61B880DEEF80E3D8E3DE45A58F324E9909060D1036441F71CE8A154EA8C356154036538A58552B533103E567AB80E882F8EA84D370561D30D7EDB24578F8BDF7853E78051E128D0790AC5F3A059E8C0A3D49F515135D08F08DDC3CE62ED8CF58E5905113A553DF62C6AD63AA331140F3B654A850E3BD0F9281EA353D6D3E19551A156CF96E90CF4E8610859FF2A55A655A28AE540E188A8A638047D33A29C73D6630E7FE4CCB8109F03D96356A607C9BA999F5A19EADD82E97948CFAB141224A93BF274AA157908D933AA87E01E24EE95ED1029106D87C9134F3CEF74BACE409549DAAA40F7006AEEA503D99B0E3B79C9336E8D0365F3E6882686332239051D97CD39C8E1998E9ED8A6C31942A74258A6B4AB523B907957761AA956DEF2450F4264DA9A8E3AEA0A644FC610B25C10CC92D993A19315DF1BDEB206B19BE7BA803A399BBE04A138E237F5C087DF381A218B8FB313C7AA72B4367A3CF4D4D9F68BF5C505A53CB5CC41C2EC21D8948E7710983F1394334168FF6B10CC9DD99B204150332B90D24B8F08CE46B6AAFF5A50DB51A9E3163C62175AB56AF47ABB0E3B60F99BB5408E7FFA0EA77CCD6D735F4835CE989350A7D38E94AE75AE2AA20721EB2AA404AD870E9566A196EF85899A5D90AE6AC6676C46A8725C5086E848168FB71C932E50E72713744FE123B307E23D19965B0FC85AFD1524C8EBCCDDB156EA08CDE182C21764BA406BAB6D5682D04F7735C1086D6CD07DB5AFC922DFB9B7DC363A42D7532BD58752ABECF818F395BC5467AA16FE5C56B984BC0621034347A892B51ED47CC77AF6FB5E3DC2472FB388A94BB09B20D606481070D4AE441D5119FD0A52BEDA24A46DC1A356DBFC462591698B7F06310BC7E821C187C01021685BDD9EB0779012DA4D83E029C8C0FFD2D199B615AF1E439C9AEE0462343801AB69222492A4E685210E0A2F670BA41BEAB4A69596D5633D1C9314D722F21DB698712D0B1A62BAC44732755170206DB38DB3C6E8385B2D51201EC40F373E82ACD6C0A3537E6BDD429C9542C4D5638A0BF199CE2097C1FF4D6A2065BAA04468B841A9D8061F8B71073CDAE12D3D816F8BA9D38A98F62715BFD511552CBC5BED7815C2E8B5D02C20AA61412A56C8C870D1707D02E52139F842A9704C960DC919D670A4C8981DE48DA7D3CA3D1499A41282EE5DD08C663417C83E356DE37DF0431F1C79B6436D1ED38AD4DAD642BEE3C83BEED297040126438439353D1533935E4426BECC428F8CA9A32094E44E45CE91FFA2F6B1329B69F39E2BBB99AA4182D39088548822CB89107142125613A9D83AC242D18A567430F1281C0021F8C9201A6460199A1B26DB38116C70034E6A321114BD49D0A83112800B01E3E16CD951A96471205063E8411CFAD2A8505122E4A48822D5C18EA3B8A57E6C7C1CDB4C33460C421421407293130F269FDB61A48646FCD80D7B26C3999AAF1E2B4CC8EB3CDAD3471A84A7F5C8675438BA1052AE7522F628532AD7A9CB5046E8206A8D4AA30C92D1A6E27144E610969AF6B6D2801A885A23F554FDBC86CEFC19B21E9D82E0CBFCD74B5C9E9583638DCA0811221A049214875035DD3DA089D49A720921ECD450FFF892B977C0C339CECBA37C1826105FFF1DCD1CC6B9880CC3D28E0CA5E3A765B5D80F7FC75AAFA66330C2F4E4F116A3D3CB3616A7803A22BD40B6D7827C277B06B98778FCA22E5661477A493DED4132FB5C83301655BCC15B7C082B106CF47357BA6DA775B133388DD6F2AB94521CE3DE71CAE6FA37BF7D690758E0612F74487767169CC7BBE60135031B6F6B3F9467652DA990F37AA840EA5D9341CAD13E9552E46BDFDCE07D8F3BDF4AFEF7966675ECD7D435D9A9CDAD731EDAAC6917C9559231E41EA7FC9652F8A9476C1224C3DA606F455446B682C0B82F07A5E95350E542149B13AD0379A7698A3C268D2E0682F19A9ADFE661C68598F1C03EA21C5FD4F10ED960A34EDE48B3379E4890729C34C33E71BA2A2ABD984CCDA0C32849C62F1C4B46206442D9C928B3B28E1372AE18FFD0BAB9790B34DFC1BBDB483323E7204E611983949B5209C6A03AA765AFA7E482D823CB32448B9F790968036AF45396B651127D0ADE896488412CB18791DE42E78BC00B2CBB71496BAE0596264964D7BA49B5FDD6A89012C2FAD8A4992C3AD86A61643B3B3109648782A1C960BE48F2D9D8B64A095556AA6C7B1B2882368D27BF4DEEA07C7A6ACB2EB7BA97F269DF020A1EC25EB7BC75A2A3C61D45C0489AB7BE7F1210003B00"
        }
        
        Add-TervisRMSCustomButton @GivexBalanceCustomPOSButtonParameters
    }    
}

function Add-GivexAdminCustomPOSButton {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$DatabaseName
    )
    process {
        Write-Verbose "$ComputerName - Adding Givex Admin button"
        $GivexAdminCustomPOSButtonParameters = @{
            ComputerName = $ComputerName
            DatabaseName = $DatabaseName
            Number = 5
            Style = 5
            Command = "http://store.givex.com"
            Description = "Givex Admin Button"
            Picture = "0x47494638396120012D00F70000FFFFFFFFFFCCFFFF99FFFF66FFFF33FFFF00FFCCFFFFCCCCFFCC99FFCC66FFCC33FFCC00FF99FFFF99CCFF9999FF9966FF9933FF9900FF66FFFF66CCFF6699FF6666FF6633FF6600FF33FFFF33CCFF3399FF3366FF3333FF3300FF00FFFF00CCFF0099FF0066FF0033FF0000CCFFFFCCFFCCCCFF99CCFF66CCFF33CCFF00CCCCFFCCCCCCCCCC99CCCC66CCCC33CCCC00CC99FFCC99CCCC9999CC9966CC9933CC9900CC66FFCC66CCCC6699CC6666CC6633CC6600CC33FFCC33CCCC3399CC3366CC3333CC3300CC00FFCC00CCCC0099CC0066CC0033CC000099FFFF99FFCC99FF9999FF6699FF3399FF0099CCFF99CCCC99CC9999CC6699CC3399CC009999FF9999CC9999999999669999339999009966FF9966CC9966999966669966339966009933FF9933CC9933999933669933339933009900FF9900CC99009999006699003399000066FFFF66FFCC66FF9966FF6666FF3366FF0066CCFF66CCCC66CC9966CC6666CC3366CC006699FF6699CC6699996699666699336699006666FF6666CC6666996666666666336666006633FF6633CC6633996633666633336633006600FF6600CC66009966006666003366000033FFFF33FFCC33FF9933FF6633FF3333FF0033CCFF33CCCC33CC9933CC6633CC3333CC003399FF3399CC3399993399663399333399003366FF3366CC3366993366663366333366003333FF3333CC3333993333663333333333003300FF3300CC33009933006633003333000000FFFF00FFCC00FF9900FF6600FF3300FF0000CCFF00CCCC00CC9900CC6600CC3300CC000099FF0099CC0099990099660099330099000066FF0066CC0066990066660066330066000033FF0033CC0033990033660033330033000000FF0000CC000099000066000033000000C0C2B8C0C3B0C3D93FC3D747C3D64FC2D457C2D35FC2D167C2D070C2CE78C2CD7FC1CB87C1C98FC1C898C1C6A0C1C5A85F5C5C5855556C69696562627F7D7D7977777270708C8A8A868484A6A5A59F9E9E999898929191C0C0C0B9B9B9B3B3B3ACACACFFFFFF00000000000000000000000000000000000021F904010000F9002C0000000020012D000008FF00EB091C48B0A0C18308132A5CC8B0A1C38710234A9C48B1A2C58B18336ADCC8B1A3C78F20438A1C49B2A4C9932853AA5CC9B2A5CB973063CA9C49B3A6CD9B3873EADCC9B3A7CF9F4079DE8B17B4A8D1A313D9A54B876E1ED2A728E7BD9BFAEE1E548DF3962E4577B5AB48775AD311F56AF15D58B164D36E04AB75AC5A8966C35A7D4B9722DBA56EEB3AB4874E2B3CBD8021DE451B782FBD774E0B2B563838AFE2797F174B8ED898AE3DA9EEDCB57B27AF5ED674ECEC09C417AFB469A2F64EC79B7B30F569D104F1BDCB9C79333E84E0BC7913872DA1B9DFBFB3891B4EAE5CBD72E6260FBC271531EB829515A6A62AEF79C378EFE8591F684F1ED579B72DDAFF537A766B5875A2076B15B8EE7C42F25B09C6531FD65DF881E6B4E9D7B6EDDC4170FB69D30D36E674634E38DCD4C38D380CE1D31C67DB4987DD54D55134213DF711C4DC54F34428D078E5B9E3563CE585E5583DF7C0775E6206DDE58E3D7745261F7DE9ACF30E6C11E1D35789253A45A3407169E5E18E4BBD33103D3C9EC5623DDF04C89F7F048D13E036BDD5E34D3DE274434E370C22740F3C44AEC8508A25AAB3247426DAF34E98E9D837103EF48566D03DEDF168643D2426E958903CBAB39D8B83650862924C9DB8D03D6C26A94E783FA278163D077D26A44048121A5667F564E3A47EDD1044CE94505A89A594E37469103D899EA58E87F5F059A29F070DFF268F3A3CA2739BA4E5AD631D8D61419A278F79A998A4AD689AA7D53A04D943ABA54B9DB990B0EBC0F30E3C756AC5CE40F0B8B3EC7A02B51316B20679AB553B02FDAA953BF4C443CFB65B8986CD39E56CE36497E7C8BB5F720391538F39C265836F41C20E9BE1400197486CB1C6D68AEBABB1594AACB9E5B9050FB35AA1F31C7DEAD84890B8DF6696E8C009DD73D6B50409FB1CC403C973D6C0225FCA5E58E838B6F09D02D53B6536D8D8AB9FBE114D4C71BB05F94CB1C508FFCCEC7D7CAE5A8F3CED1DEC1ABBE9C6031BC4E85425903C1C2FE50E41EA415A90CA61B5F3DC3CD5A6A38E43956A7D50D94BA23C50D93252FAEDD567615A109FE00E246580DF74FF13A0A90F410CDA58DDB9E818D5562F9D759B452F058F55F72CBED58DF5C07916CD7CCA0823C8D10DC4AED2052DECD66079135436CDC9B2EBEC417CA28E6D58A8BB0DE4DCA6F72A908A67B7769675DE6CBA1F38130D7670E8EB18FAF976A2737D79E3E9085A36C94B5F8E63AC26361C16C8B7873DD060718F565F42F868CF50EBACC32EDF591A9EE556F861C1F6BC6AA795ED587EBE5F2911FB6DE5B832422A92BB3DED045918F45E77AEDAC18C72D4CBDF40F834C082B48C5BF5188CD708C2A7CD50E58253095343187810A89D2F2C05C91AF454043DA3A9CF200072D28026C2A7AD4584835E42DFFFFC6210FCE1C520B87261B90C3641E529B05B6159DDDBFFAE2790CE0D44723F63D5402026B3FD2D51862983D987C244381356AF209A725295249235212AA48B0A29DB7D06B3BADD19C45C3A14C83C52F5A2A2E5C5882DBA221C2368C51F86F12CF0288D3C8495463C4171885AE90CAE7257B93ADEB020360B10CF2432478634728687A4A31D07F2C72716D08134226411AFB8C949BAB1939184A4150D75C69F0D6F8720BC5B58AE4546828C0C83B0A4CAF432A5B3FD2428789C7CC82341491849869292A92C081A11329FF2F4D0882AEAA141C2349739664D1EF08B6669669990353ACC59B25BCECA6036BD30B9CE21D8F09BEFC69114DB4524990A61A6280D554954AA2D21F20893268D9836C611538ABC3C119FECC6FF91829D2B42D93CE279C2D2BDAC950E9CE2F41D952452CF3E36A4A109311757D6799076FAF192DC11E61F8DF8C0A5F09320EC82DE1C6DA8498DB06B1DB4618776226A51C14D2A80238BE87F02E48D4FF14D221D4DC7479375909CEE542021F5E12F05D2D2EF71AD1D38B2C74673E9CB4215C41E8B634D2307D3C0D1A8E3710891CA89EAE93F870414906771285FCE22B6AFA9031DB34C21A77AD3BB009133ABEF3891F09C351E74FC54788E81AA19296A90A262B41EE441471E6705C0A63A0651649D4769C08447A1F6F2835A31D3D448689D6D756F302232CD54D205508BD6A39E5A591DD854C58E6C11097536D54F7F6AB6A96F90831CD900290DAD579FA8FFC983B14B41AA03D9D40EC5C603B7B3E56B41FCFA4EC0320B75735C18A1D0E358520A8D62FC5CD8739058A68F7EF543B97A8F2961935AFDFC0B4BBEBBA51AF73A10E512CAA1E6E511739B5BD16042D69EC625D43AA6D7C8F4AA6A96BBF46779CEE4AABCD8B74403BC6EF6CC971057D5CA2A89DC9941DAEA560A9E5095145B6FE8224CCD461217BEAD4AD23A06F6C87894AD3CF0A830530B822AF99E48A2B3D46F8946E4D98BBE34A2BC620AE512AC0DE019041B4DDACF422D093483C4D352218EE887F1484DC3B61782EFED638AC264236A668B36D8FBD03CD85523AC1EE4C999893277DE41E536CDA3C8F5F059530EF25BDA98D91DBCE58E6A14721A2D3B90931E5866475C09E22FE07CD720E610C737C471E7306FC58BF6A0077DA2A544295399D009C1B29B5A64E683C08936DD7B53690A1D11D3801923A671F372A42611BDFE952C43B9744148531A0B4D533939010F9BF189EA562F066C087C2A665D4DEBC0E8286C712D0D64D8F4CD5AFBFA2A6335215A7F4D6CAF284BD89A2EB6B27D42368AC16AD9D0860A3E4C1B2278243BDA3B0908003B00"
        }
        
        Add-TervisRMSCustomButton @GivexAdminCustomPOSButtonParameters        
    }    
}

function Set-GivexRMSItemProperties {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$DatabaseName
    ) 
    begin {
        $SetItemPropertiesQuery = @"
UPDATE Item
SET Taxable = 0, PriceMustBeEntered = 1, LastUpdated = GETDATE()
WHERE ItemLookupCode = 'GIVEXACT'       
"@
    }
    process {
        Write-Verbose "$ComputerName - Setting Givex RMS item properties"
        Invoke-RMSSQL -DataBaseName $DatabaseName -SQLServerName $ComputerName -Query $SetItemPropertiesQuery
    }    
}

function Install-GivexGcmIniFile_DEV {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    begin {
        $UserCredentials = Get-PasswordstatePassword -ID 5450
        $OperatorCredentials = Get-PasswordstatePassword -ID 5451
        $GcmIniLocalPath = "C:\Program Files\Microsoft Retail Management System\Store Operations\gcm.ini"
        $TemplateVariables = @{
            UserID = $UserCredentials.UserName
            UserPassword = $UserCredentials.Password
            OperatorID = $OperatorCredentials.UserName
            OperatorPassword = $OperatorCredentials.Password
            URL = $UserCredentials.URL.split(":")[0]
            Port = $UserCredentials.URL.split(":")[1]
        }
    }
    process {
        $GcmIniContent = Invoke-ProcessTemplateFile -TemplateFile $PSScriptRoot\Templates\gcm.ini.pstemplate -TemplateVariables $TemplateVariables
        $RemoteGcmIniPath = $GcmIniLocalPath | ConvertTo-RemotePath -ComputerName $ComputerName
        $GcmIniContent | Out-File -FilePath $RemoteGcmIniPath -Force -Encoding utf8
    }    
}

function Install-GivexGcmIniFile {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    begin {
        $GcmIniLocalPath = "C:\Program Files\Microsoft Retail Management System\Store Operations\gcm.ini"
        $GivexStoreCredentialTable = Get-GivexStoreCredentialTable
    }
    process {
        Write-Verbose "$ComputerName - Installing GCM.ini file"
        $StoreCredential = Get-GivexStoreCredential -GivexStoreCredentialTable $GivexStoreCredentialTable -ComputerName $ComputerName

        $TemplateVariables = @{
            UserID = $StoreCredential."User ID"
            UserPassword = $StoreCredential."User Password"
        }

        $GcmIniContent = Invoke-ProcessTemplateFile -TemplateFile $PSScriptRoot\Templates\gcm.ini.pstemplate -TemplateVariables $TemplateVariables
        $RemoteGcmIniPath = $GcmIniLocalPath | ConvertTo-RemotePath -ComputerName $ComputerName
        $GcmIniContent | Out-File -FilePath $RemoteGcmIniPath -Force -Encoding utf8
    }    
}

function Get-GivexStoreCredentialTable {
    $StoreInfoObject =  Get-PasswordstatePassword -ID 5526
    $StoreInfoObject.GenericField1 | ConvertFrom-Json
}

function Get-GivexStoreCredential {
    param (
        $GivexStoreCredentialTable = (Get-GivexStoreCredentialTable),
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    process {
        $StoreCode = $ComputerName.Substring(0,4)
        $RegisterNumber = (($ComputerName -split "POS")[1] -split "-")[0]
        $GivexStoreCredential = $GivexStoreCredentialTable | 
            Where-Object {
                ($_.StoreCode -eq $StoreCode) -and
                ($_."User Description" -match "POS$RegisterNumber")
            }
        if ($GivexStoreCredential) {
            $GivexStoreCredential
        } else {
            Write-Warning "$ComputerName - No Givex credential found"
        }
    }
}

function Invoke-nChannelSyncManagerProvision {
    param (
        $EnvironmentName = "Delta"
    )
    Invoke-ApplicationProvision -ApplicationName nChannelSyncManager -EnvironmentName $EnvironmentName
    #$Nodes = Get-TervisApplicationNode -ApplicationName nChannelSyncManager -EnvironmentName $EnvironmentName
}

function Invoke-StoreExchangeMailboxToMigaduMailboxMigration {
    $TervisStoreDefinition = Get-TervisStoreDefinition -Name "New Orleans"
    
    get-service -ComputerName exchange2016 -Name msExchangeIMAP4 | fl *
    get-service -ComputerName exchange2016 -Name msExchangeIMAP4 | start-service

    get-service -ComputerName exchange2016 -Name MSExchangeIMAP4BE | start-service

    $TervisStoreDefinition.BackOfficeADUser.samaccountname
    $TervisStoreDefinition.BackOfficeUserCredential.GetNetworkCredential().password

    get-ExchangeCASMailbox -Identity neworleansstore

    Test-ExchangeImapConnectivity -MailboxCredential $TervisStoreDefinition.BackOfficeUserCredential
    $Credential = New-Crednetial -Username $TervisStoreDefinition.BackOfficeADUser.UserPrincipalName -Password $TervisStoreDefinition.BackOfficeUserCredential.GetNetworkCredential().password
    $Result = Test-ExchangeImapConnectivity -MailboxCredential $TervisStoreDefinition.BackOfficeUserCredential
    $Result = Test-ExchangeImapConnectivity -MailboxCredential $Credential
    

    $TervisStoreDefinition.MigaduMailboxCredential.UserName
    $TervisStoreDefinition.MigaduMailboxCredential.GetNetworkCredential().password
    $MigaduEmailServerConfiguration = [PSCustomObject]@{
        IMAPServerName = "imap.migadu.com"
        IMAPPort = 993
        IMAPConnectionSecurity = "SSL/TLS"
        SMTPServerName = "smtp.migadu.com"
        SMTPPort = 587
        SMTPConnectionSecurity = "STARTTLS"
    }

    "imapsync --host1 exchange2016.tervis.prv --host2 $($MigaduEmailServerConfiguration.IMAPServerName) --justconnect"
    "imapsync --host1 exchange2016.tervis.prv --exchange1 --user1 $($TervisStoreDefinition.BackOfficeADUser.UserPrincipalName) --password1 $($TervisStoreDefinition.BackOfficeUserCredential.GetNetworkCredential().password) --host2 $($MigaduEmailServerConfiguration.IMAPServerName) --user2 $($TervisStoreDefinition.MigaduMailboxCredential.UserName) --password2 $($TervisStoreDefinition.MigaduMailboxCredential.GetNetworkCredential().password) --justconnect"
    "imapsync --host1 exchange2016.tervis.prv --exchange1 --user1 $($TervisStoreDefinition.BackOfficeADUser.UserPrincipalName) --password1 $($TervisStoreDefinition.BackOfficeUserCredential.GetNetworkCredential().password) --host2 $($MigaduEmailServerConfiguration.IMAPServerName) --user2 $($TervisStoreDefinition.MigaduMailboxCredential.UserName) --password2 $($TervisStoreDefinition.MigaduMailboxCredential.GetNetworkCredential().password) --justlogin"
    "imapsync --host1 exchange2016.tervis.prv --exchange1 --user1 $($TervisStoreDefinition.BackOfficeADUser.UserPrincipalName) --password1 $($TervisStoreDefinition.BackOfficeUserCredential.GetNetworkCredential().password) --host2 $($MigaduEmailServerConfiguration.IMAPServerName) --user2 $($TervisStoreDefinition.MigaduMailboxCredential.UserName) --password2 $($TervisStoreDefinition.MigaduMailboxCredential.GetNetworkCredential().password) --justfoldersizes"
    "imapsync --host1 exchange2016.tervis.prv --exchange1 --user1 $($TervisStoreDefinition.BackOfficeADUser.UserPrincipalName) --password1 $($TervisStoreDefinition.BackOfficeUserCredential.GetNetworkCredential().password) --host2 $($MigaduEmailServerConfiguration.IMAPServerName) --user2 $($TervisStoreDefinition.MigaduMailboxCredential.UserName) --password2 $($TervisStoreDefinition.MigaduMailboxCredential.GetNetworkCredential().password) --justfoldersizes --automap"
    "imapsync --host1 exchange2016.tervis.prv --exchange1 --user1 $($TervisStoreDefinition.BackOfficeADUser.UserPrincipalName) --password1 $($TervisStoreDefinition.BackOfficeUserCredential.GetNetworkCredential().password) --host2 $($MigaduEmailServerConfiguration.IMAPServerName) --user2 $($TervisStoreDefinition.MigaduMailboxCredential.UserName) --password2 $($TervisStoreDefinition.MigaduMailboxCredential.GetNetworkCredential().password) --automap"

    
}