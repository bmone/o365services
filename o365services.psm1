# Connect to Office 365 Services
# Tomasz Knop | Tuesday 25 Feb 2020

function Connect-O365Service {
    <#
    .SYNOPSIS
    .DESCRIPTION
    #>
    [CmdletBinding()]
    param(
        [parameter()] [validateSet("AzureAD", "AzureADLegacy", "ExchangeOnPrem", "ExchangeOnline", "SecurityAndCompliance", "SharePoint", "Skype", "Teams", "All")] [string[]] $Service = @("All"),
        [parameter(Mandatory)] [validateNotNullorEmpty()] [string] $UserPrincipalName,
        [parameter()] [validateNotNullorEmpty()] [string] $TenantID, 
        [parameter()] [switch] $MFA,
        [parameter()] [switch] $Disconnect
    )
    begin {
        Write-Verbose -Message "[BEGIN] Starting $($MyInvocation.MyCommand)"
        $moduleVerbose = $false
        $o365Services = @( "AzureAD", "AzureADLegacy", "ExchangeOnPrem", "ExchangeOnline", "SecurtiyAndCompliance", "SharePoint", "Skype", "Teams" )
        $connectO365Services = [ordered]@{ }
        $o365Services.ForEach( {
                $tryService = if ($Service -contains "All") { $true } else { $Service -contains $PSItem } 
                $connectO365Services[$PSItem] = $tryService 
            }
        )
        $tryO365Services = ($connectO365Services.GetEnumerator() | Where-Object Value -eq $true).Name
        Write-Verbose -Message ("[BEGIN] Using UserPrincipalName '{0}'." -f $UserPrincipalName)
        $userCredential = $null
        # Use Credentials only if MFA is not required
        if (-not $MFA.IsPresent) {
            # Look up stored Secret if available
            if ($null -eq (Get-Module -Name "Microsoft.PowerShell.SecretsManagement" -ListAvailable -Verbose:$moduleVerbose)) {
                Write-Verbose -Message ("[BEGIN] Microsoft.PowerShell.SecretsManagement Module is not present." -f $service)
            }
            else {
                Import-Module -Name "Microsoft.PowerShell.SecretsManagement" -Verbose:$moduleVerbose
                Write-Verbose -Message "[BEGIN] Looking up stored secrets."
                if ($cachedCredentials = Get-SecretInfo | Where-Object TypeName -eq PSCredential | 
                    ForEach-Object { Get-Secret $PSItem.Name -Vault $PSItem.Vault | Where-Object UserName -eq $UserPrincipalName } | 
                    Select-Object -First 1) {
                    Write-Verbose -Message "[BEGIN] Secret found and retreived."
                    $userCredential = $cachedCredentials
                }
            }
            # Fallback to standard Credential prompt
            if (-not $userCredential) {
                Write-Verbose -Message "[BEGIN] No stored Secret found. Prompting ..."
                $userCredential = Get-Credential -Message "Enter Credential to use (UPN)" -UserName $UserPrincipalName               
            }
        }
        Write-Output "Selected services: $($tryO365Services -join ' | ')"
    }
    process {
        $connectedServices = @()
        switch ( $tryO365Services ) {
            "AzureADLegacy" {
                $o365Service = $PSItem
                # AzureAD v1 Legacy [MSOnline] (skip if tenantID supplied)
                if (-not $TenantID) {
                    if ($null -eq (Get-Module -Name "MSOnline" -ListAvailable -Verbose:$moduleVerbose)) {
                        Write-Warning -Message ("[PROCESS] Skipping service '{0}'. MSOnline Module is not installed or present." -f $o365Service)
                    }
                    else {
                        $paramAzureAD = @{ Credential = $userCredential }
                        if ($MFA.IsPresent) { $paramAzureAD.Remove('Credential') }
                        Import-Module MSOnline -Verbose:$moduleVerbose
                        try {
                            Write-Verbose -Message ("[PROCESS] [TRY] Attempting to connect to service '{0}' ..." -f $o365Service)
                            $null = Connect-MsolService @paramAzureAD -ErrorAction Stop
                            if ($tID = Get-MsolCompanyInformation) { $o365Service += (" [{0}]" -f $tID.DisplayName) }
                            Write-Verbose -Message ("[PROCESS] [TRY] {0} connected." -f $o365Service)
                            $connectedServices += $o365Service
                        }
                        catch {
                            Write-Warning -Message ("[PROCESS] [CATCH] Service '{0}' - Connection exception occured!" -f $o365Service)
                            $PSCmdlet.ThrowTerminatingError($PSitem)
                        }
                    }
                }
            }
            "AzureAD" {
                # AzureAD v2 Modern [AzureAD]
                $o365Service = $PSItem
                if ($null -eq (Get-Module -Name "AzureAD" -ListAvailable -Verbose:$moduleVerbose)) {
                    Write-Warning -Message ("[PROCESS] Skipping {0}! AzureAD Module is not present." -f $o365Service)
                }
                else {
                    $paramAzureAD = @{ Credential = $userCredential }
                    switch ($true) {
                        { $MFA.IsPresent } { $paramAzureAD.Add('AccountId', $UserPrincipalName) ; $paramAzureAD.Remove('Credential') }
                        { $TenantID } { $paramAzureAD.Add('TenantID', $TenantID) ; }
                    }
                    Import-Module AzureAD -Verbose:$moduleVerbose
                    try { 
                        Write-Verbose -Message ("[PROCESS] [TRY] Attepmting to connect to {0} ..." -f $o365Service)
                        $null = Connect-AzureAD @paramAzureAD -ErrorAction Stop
                        if ($tID = Get-AzureADTenantDetail) { $o365Service += (" [{0}]" -f $tID.DisplayName) }
                        Write-Verbose -Message ("[PROCESS] [TRY] {0} connected." -f $o365Service)
                        $connectedServices += $o365Service
                    }
                    catch {
                        Write-Warning -Message ("[PROCESS] [CATCH] {0} - Connection exception occured!" -f $o365Service)
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
