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

    $Responses = Start-ParallelWork -ScriptBlock {
        param($Parameter)
        Install-TervisOffice2016VLPush -ComputerName $Parameter
    } -Parameters $BackOfficeComputers
}