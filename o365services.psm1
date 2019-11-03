#Requires -Modules CredStore 
# Office 365 Services Functions
# Tomasz Knop | 9 October 2019

function Connect-O365Service {
    <#
    .SYNOPSIS
    .DESCRIPTION
    #>
    [CmdletBinding()]
    param(
        [parameter()] [validateSet("AzureAD", "ExchangeOnPrem", "Exchange", "SharePoint", "SecAndCompl", "Skype", "Teams", "All")] [string[]] $Services = @("All"),
        [parameter()] [string] $UserPrincipalName = [System.Environment]::GetEnvironmentVariable("UPN"),
        [parameter()] [validateNotNullorEmpty()] [string] $TenantID, 
        [parameter()] [switch] $MFA,
        [parameter()] [switch] $Disconnect
    )
    begin {
        $moduleVerbose = $false
        Write-Verbose -Message "[BEGIN] Starting $($MyInvocation.Mycommand)"
        $availableServices = @("AzureAD", "ExchangeOnPrem", "Exchange", "SharePoint", "SecAndCompl", "Skype", "Teams")
        $connectServices = [ordered]@{ }
        $tryServices = @()
        $availableServices.ForEach( {
                $tryService = if ($Services -contains "All") { $true } else { $Services -contains $PSItem }
                $connectServices.Add($PSItem, $tryService)
                if ($tryService) { $tryServices += $PSItem }
            })
        $connectedServicesList = @()
        Write-Output "UserPrincipalName: $UserPrincipalName"
        $userCredential = Get-MyCredential -Username $UserPrincipalName
    }
    process {
        switch ( $tryServices ) {
            "AzureAD" {
                # AzureAD v1 (skip if tenantID supplied)
                if (-not $TenantID) {
                    $service = "AzureAD (v1)"
                    if ($null -eq (Get-Module -Name "MSOnline" -ListAvailable -Verbose:$moduleVerbose)) {
                        Write-Warning -Message ("[PROCESS] MSOnline Module is not present! Skipping {0}" -f $service)
                        continue
                    }
                    else {
                        $paramAzureAD = @{ Credential = $userCredential }
                        if ($MFA.IsPresent) { $paramAzureAD.Remove('Credential') }
                        Import-Module MSOnline -Verbose:$moduleVerbose
                        try {
                            Write-Verbose -Message ("[PROCESS][TRY] Connecting to $service")
                            $null = Connect-MsolService @paramAzureAD -ErrorAction Stop
                            if ($tID = Get-MsolCompanyInformation) { $service += (" [{0}]" -f $tID.DisplayName) }
                            Write-Verbose -Message ("[PROCESS][TRY] $service connected.")
                            $connectedServicesList += $service
                        }
                        catch {
                            Write-Warning -Message ("[PROCESS][CATCH] {0} - Connection exception occured!" -f $service)
                            $PSCmdlet.ThrowTerminatingError($PSitem)
                        }
                    }
                }
                # AzureAD v2
                $service = "AzureAD (v2)"
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
                        Write-Verbose -Message ("[PROCESS][TRY] Connecting to $service")
                        $null = Connect-AzureAD @paramAzureAD -ErrorAction Stop
                        if ($tID = Get-AzureADTenantDetail) { $service += (" [{0}]" -f $tID.DisplayName) }
                        Write-Verbose -Message ("[PROCESS][TRY] $service connected.")
                        $connectedServicesList += $service
                    }
                    catch {
                        Write-Warning -Message ("[PROCESS][CATCH] {0} - Connection exception occured!" -f $service)
                        $PSCmdlet.ThrowTerminatingError($PSitem)
                    }
                }
            }
            "ExchangeOnPrem" { }
            "Exchange" { }
            "SharePoint" { }
            "SecAndCompl" { }
            "Skype" { }
            "Teams" { }
        }
    }
    end {
        Write-Output "Connected services: $($connectedServicesList -join ' | ')"
        Write-Verbose -Message "[END] Ending $($MyInvocation.Mycommand)"
    }
}