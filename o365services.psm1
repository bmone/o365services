# Connect to Office 365 Services
# Tomasz Knop | Tuesday 25 Feb 2020

function Connect-O365Service {
    <#
    .SYNOPSIS
    .DESCRIPTION
    #>
    [CmdletBinding()]
    param(
        [parameter()] [validateSet("AzureAD", "ExchangeOnPrem", "ExchangeOnline", "SharePoint", "SecurityAndCompliance", "Skype", "Teams", "All")] [string[]] $Services = @("All"),
        [parameter(Mandatory)] [validateNotNullorEmpty()] [string] $UserPrincipalName,
        [parameter()] [validateNotNullorEmpty()] [string] $TenantID, 
        [parameter()] [switch] $MFA,
        [parameter()] [switch] $Legacy,
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
        $userCredential = $null
        Write-Verbose -Message ("[BEGIN] Using UserPrincipalName '{0}'." -f $UserPrincipalName)
        # Use Credentials only if MFA is not required
        if (-not $MFA.IsPresent) {
            # Look up stored Secret if available
            if ($null -eq (Get-Module -Name "Microsoft.PowerShell.SecretsManagement" -ListAvailable -Verbose:$moduleVerbose)) {
                Write-Verbose -Message ("[PROCESS] Microsoft.PowerShell.SecretsManagement Module is not present." -f $service)
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
    }
    process {
        Write-Output "Selected services: $(($connectServices.GetEnumerator() | Where-Object Value -eq $true).Name -join ' | ')"
        $connectedServices = @()
        switch ( $tryServices ) {
            "AzureAD" {
                # AzureAD v1 Legacy [MSOnline] (skip if tenantID supplied)
                if ( (-not $TenantID) -and ($Legacy.IsPresent) ) {
                    $service = "AzureAD (Legacy)"
                    if ($null -eq (Get-Module -Name "MSOnline" -ListAvailable -Verbose:$moduleVerbose)) {
                        Write-Verbose -Message ("[PROCESS] MSOnline Module is not present! Skipping {0}" -f $service)
                    }
                    else {
                        $paramAzureAD = @{ Credential = $userCredential }
                        if ($MFA.IsPresent) { $paramAzureAD.Remove('Credential') }
                        Import-Module MSOnline -Verbose:$moduleVerbose
                        try {
                            Write-Verbose -Message ("[PROCESS] [TRY] Connecting to $service ...")
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
                        Write-Verbose -Message ("[PROCESS] [TRY] Connecting to $service ...")
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
