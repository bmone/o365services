# Connect to Office 365 Services
# Tomasz Knop | Tuesday 25 Feb 2020

function Read-PromptChoice {
    [CmdletBinding()]
    param( 
        [parameter()] [string] $Title = [string]::Empty,
        [parameter()] [string] $Message = "Make your choice",
        [parameter()] [string[]] $Choices = @("&Yes#Choose Yes", "&No#Choose No"),
        [parameter()] [uint16] $DefaultChoice = 0
    )
    # $Choices example : "&Yes#Delete all","&No#Abort operation"
    [System.Management.Automation.Host.ChoiceDescription[]] $Options = $Choices |
        ForEach-Object { New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList $PSItem.Split('#') }
    return ($Host.UI.PromptForChoice($Title, $Message, $Options, $DefaultChoice))
}

function Connect-O365Service {
    <#
    .SYNOPSIS
    .DESCRIPTION
    #>
    [CmdletBinding()]
    param(
        [parameter()] [validateSet("AzureAD", "ExchangeOnPrem", "ExchangeOnline", "SharePoint", "SecurityAndCompliance", "Skype", "Teams", "All")] [string[]] $Services = @("All"),
        [parameter()] [string] $UserPrincipalName = [System.Environment]::GetEnvironmentVariable("UPN"),
        [parameter()] [validateNotNullorEmpty()] [string] $TenantID, 
        [parameter()] [switch] $MFA,
        [parameter()] [switch] $Disconnect
    )
    begin {
        Write-Verbose -Message "[BEGIN] Starting $($MyInvocation.MyCommand)"
        $availableServices = @("AzureAD", "ExchangeOnPrem", "ExchangeOnline", "SharePoint", "SecurtiyAndCompliance", "Skype", "Teams")
        $connectServices = [ordered]@{}
        $tryServices = @()
        $availableServices.ForEach( {
                $tryService = if ($Services -contains "All") { $true } else { $Services -contains $PSItem }
                $connectServices.Add($PSItem, $tryService)
                if ($tryService) { $tryServices += $PSItem }
            })
        $moduleVerbose = $false
        Write-Verbose -Message ("[BEGIN] Using UserPrincipalName '{0}'." -f $UserPrincipalName)
        $cachedCredentials = @{}
        if ($null -eq (Get-Module -Name "Microsoft.PowerShell.SecretsManagement" -ListAvailable -Verbose:$moduleVerbose)) {
            Write-Verbose -Message ("[PROCESS] Microsoft.PowerShell.SecretsManagement Module is not present." -f $service)
            continue
        }
        else {
            Import-Module -Name "Microsoft.PowerShell.SecretsManagement" -Verbose:$moduleVerbose
            Write-Verbose -Message ("[BEGIN] Looking up UserPrincipalName in Credential vaults." -f $UserPrincipalName) 
            foreach ( $item in (Get-SecretInfo | Where-Object TypeName -eq PSCredential) ) { 
                $getCredential = Get-Secret $item.Name -Vault $item.Vault
                if ($getCredential.UserName -eq $UserPrincipalName) {
                        $cachedCredentials[$item.Name] = $getCredential
                }
            }
        }
        switch ($cachedCredentials.Count) {
            0 {
                Write-Verbose -Message "[BEGIN] No cached Credentials found. Prompting ..."
                $userCredential = Get-Credential -Message "Enter Credential to use (UPN)" -UserName $UserPrincipalName               
            }
            1 {
                Write-Verbose -Message ("[BEGIN] Found cached Credential for '{0}'." -f $UserPrincipalName)
                $userCredential = $cachedCredentials
            }
            {$PSItem -gt 1} {
                Write-Verbose -Message "[BEGIN] Found many cached Credentials. Prompting ..."
                $userCredential = $cachedCredentials[(Read-PromptChoice -Message "Which cached Credential would you like to use for $($UserPrincipalName)?" -Choices $cachedCredentials.Keys)] 
            }
        }
        Write-Output "Requested services: $tryServices"
    }
    process {
        $connectedServices = @()
        switch ( $tryServices ) {
            "AzureAD" {
                # AzureAD v1 (skip if tenantID supplied or MFA is in use)
                if ( (-not $TenantID) -and (-not $MFA.IsPresent) ) {
                    $service = "AzureAD (legacy)"
                    if ($null -eq (Get-Module -Name "MSOnline" -ListAvailable -Verbose:$moduleVerbose)) {
                        Write-Verbose -Message ("[PROCESS] MSOnline Module is not present! Skipping {0}" -f $service)
                        continue
                    }
                    else {
                        $paramAzureAD = @{ Credential = $userCredential }
                        if ($MFA.IsPresent) { $paramAzureAD.Remove('Credential') }
                        Import-Module MSOnline -Verbose:$moduleVerbose
                        try {
                            Write-Verbose -Message ("[PROCESS] [TRY] Connecting to $service")
                            $null = Connect-MsolService @paramAzureAD -ErrorAction Stop
                            if ($tID = Get-MsolCompanyInformation) { $service += (" [{0}]" -f $tID.DisplayName) }
                            Write-Verbose -Message ("[PROCESS] [TRY] $service connected.")
                            $connectedServices += $service
                        }
                        catch {
                            Write-Warning -Message ("[PROCESS] [CATCH] {0} - Connection exception occured!" -f $service)
                            $PSCmdlet.ThrowTerminatingError($PSitem)
                        }
                    }
                }
                # AzureAD v2
                $service = "AzureAD"
                if ($null -eq (Get-Module -Name "AzureAD" -ListAvailable -Verbose:$moduleVerbose)) {
                    Write-Warning -Message ("[PROCESS] AzureAD Module is not present! Skipping {0}" -f $service)
                    continue
                }
                else {
                    $paramAzureAD = @{ Credential = $userCredential }
                    switch ($true) {
                        { $MFA.IsPresent } { $paramAzureAD.Add('AccountId', $UserPrincipalName) ; $paramAzureAD.Remove('Credential') }
                        { $TenantID } { $paramAzureAD.Add('TenantID', $TenantID) ; }
                    }
                    Import-Module AzureAD -Verbose:$moduleVerbose
                    try { 
                        Write-Verbose -Message ("[PROCESS] [TRY] Connecting to $service")
                        $null = Connect-AzureAD @paramAzureAD -ErrorAction Stop
                        if ($tID = Get-AzureADTenantDetail) { $service += (" [{0}]" -f $tID.DisplayName) }
                        Write-Verbose -Message ("[PROCESS] [TRY] $service connected.")
                        $connectedServices += $service
                    }
                    catch {
                        Write-Warning -Message ("[PROCESS] [CATCH] {0} - Connection exception occured!" -f $service)
                        $PSCmdlet.ThrowTerminatingError($PSitem)
                    }
                }
            }
            "ExchangeOnPrem" { }
            "ExchangeOnline" { }
            "SharePoint" { }
            "SecurtiyAndCompliance" { }
            "Skype" { }
            "Teams" { }
        }
    }
    end {
        Write-Output "Connected services: $($connectedServices -join ' | ')"
        Write-Verbose -Message "[END] Ending $($MyInvocation.MyCommand)"
    }
}
Export-ModuleMember -Function Connect-O365Service