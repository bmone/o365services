# Connect to Office 365 Services
# Tomasz Knop | Tuesday 25 Feb 2020

function Connect-O365Service {
    <#
    .SYNOPSIS
    .DESCRIPTION
    #>
    [CmdletBinding()]
    param(
        [parameter()] [validateSet("AzureAD", "AzureADLegacy", "ExchangeOnPrem", "ExchangeOnline", "SecurityAndCompliance", "SharePointOnline", "Teams", "InformationProtection","All")] [string[]] $Service = @("All"),
        [parameter(Mandatory,HelpMessage="You must provide a valid user account name in UPN format.")] [validateNotNullorEmpty()] [string] $UserPrincipalName,
        [parameter()] [validateNotNullorEmpty()] [string] $TenantID, 
        [parameter()] [switch] $MFA,
        [parameter()] [switch] $Disconnect
    )
    begin {
        Write-Verbose -Message "[BEGIN] Starting $($MyInvocation.MyCommand)"
        $verboseFlag = $false
        $useMFA = $PSBoundParameters.ContainsKey('MFA')
        Write-Verbose -Message ("[BEGIN] Using MFA : {0}" -f $useMFA)
        $global:connectO365Services = @{}
        $o365Services = @( "AzureAD", "AzureADLegacy", "ExchangeOnPrem", "ExchangeOnline", "SecurtiyAndCompliance", "SharePointOnline", "Teams", "InformationProtection" )
        $o365Services | ForEach-Object {
                            $tryService = if ($Service -contains "All") { $true } else { $Service -contains $PSItem }
                            if ($tryService) { $global:connectO365Services[$PSItem] = @{ NotConnected = $true } }
                        }
        $userCredential = @{}
        # Use Credentials only if MFA is not required
        if (-not $useMFA) {
            # See if SecretManagement module is installed.
            if ($null -eq (Get-Module -Name "Microsoft.PowerShell.SecretManagement" -ListAvailable -Verbose:$verboseFlag)) {
                Write-Verbose -Message "[BEGIN] Microsoft.PowerShell.SecretManagement Module is not present. Secret Vault(s) unavailable."
                # If no, fallback to standard Credential prompt
                Write-Verbose -Message "[BEGIN] Falling back to Credential Prompt ..."
                $userCredential.Credential = (Get-Credential -UserName $UserPrincipalName -Message "Enter credential to use (UPN)")
            } 
            else {
                # If yes, load it up and check if UPN matches any stored secret
                Import-Module -Name "Microsoft.PowerShell.SecretManagement" -Verbose:$verboseFlag
                Write-Verbose -Message "[BEGIN] Looking up stored credential in Secret Vault(s) ..."
                $vaultCredential = Get-SecretInfo | Where-Object Type -eq PSCredential | ForEach-Object { Get-Secret $PSItem.Name -Vault $PSItem.VaultName | Where-Object UserName -eq $UserPrincipalName }
                # If so, use it
                if ($vaultCredential) {
                    Write-Verbose -Message ("[BEGIN] Credential for '{0}' found and successfully retreived from the vault." -f $vaultCredential.UserName)
                    $userCredential.Credential = $vaultCredential
                }
            }
        }
        Write-Output ("Selected O365Services: {0}" -f (($global:connectO365Services.GetEnumerator()).Name -join ' | '))
        Write-Output ("Using UPN : {0}" -f $UserPrincipalName)
        Write-Output ("Using MFA : {0}" -f $useMFA)
    }
    process {
        switch ( ($global:connectO365Services.GetEnumerator()).Name ) {
            # AzureAD v1 Legacy [MSOnline]
            "AzureADLegacy" {
                $o365Service = $PSItem
                Write-Output ("Linking : {0}" -f $PSItem)
                # Skip if tenantID supplied
                if (-not $TenantID) {
                    if ($null -eq (Get-Module -Name "MSOnline" -ListAvailable -Verbose:$verboseFlag)) {
                        Write-Warning -Message ("Service {0} : Skipping, MSOnline module is not present." -f $o365Service)
                    }
                    else {
                        Import-Module MSOnline -Verbose:$verboseFlag
                        try {
                            Write-Verbose -Message ("[PROCESS] [TRY] Attempting to connect service {0} ..." -f $o365Service)
                            $null = Connect-MsolService @userCredential -ErrorAction Stop
                            if ($tID = Get-MsolCompanyInformation -ErrorAction SilentlyContinue) { 
                                Write-Verbose -Message ("[PROCESS] [TRY] Service {0} connected." -f $o365Service)
                                $global:connectO365Services[$o365Service]["Status"] = ("{0} [{1}] ({2})" -f $o365Service, $tID.DisplayName,(Get-Date))
                                $global:connectO365Services[$o365Service]["Connected"] = $true
                                $global:connectO365Services[$o365Service].Remove("NotConnected")
                            }
                        }
                        catch {
                            Write-Warning -Message ("[PROCESS] [CATCH] Service {0} : Connection exception occured!" -f $o365Service)
                            $PSCmdlet.ThrowTerminatingError($PSitem)
                        }
                    }
                }
            }
            # AzureAD v2 Modern [AzureAD]
            "AzureAD" {
                $o365Service = $PSItem
                Write-Output ("Linking : {0}" -f $PSItem)
                if ($null -eq (Get-Module -Name "AzureAD" -ListAvailable -Verbose:$verboseFlag)) {
                    Write-Warning -Message ("Service {0} : Skipping, AzureAD Module is not present." -f $o365Service)
                }
                else {
                    $params = @{}
                    switch ($true) {
                        $useMFA { $params.AccountId = $UserPrincipalName }
                        $TenantID { $params.TenantID = $TenantID }
                    }
                    Import-Module AzureAD -Verbose:$verboseFlag
                    try { 
                        Write-Verbose -Message ("[PROCESS] [TRY] Attepmting to connect service {0} ..." -f $o365Service)
                        $null = Connect-AzureAD @params @userCredential -ErrorAction Stop
                        if ($tID = Get-AzureADTenantDetail -ErrorAction SilentlyContinue) {
                            Write-Verbose -Message ("[PROCESS] [TRY] Service {0} connected." -f $o365Service)
                            $global:connectO365Services[$o365Service]["Status"] = ("{0} [{1}|{2}] ({3})" -f $o365Service,$tID.DisplayName,$tID.ObjectId,(Get-Date))
                            $global:connectO365Services[$o365Service]["Connected"] = $true
                            $global:connectO365Services[$o365Service].Remove("NotConnected")
                        }   
                    }
                    catch {
                        Write-Warning -Message ("[PROCESS] [CATCH] Service {0} : Connection exception occured!" -f $o365Service)
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
            "InformationProtection" {
#                # Azure Information Protection [AIP]
#                $o365Service = $o365ServiceTemplate.Clone()
#                $o365Service.Name = $PSItem
#                $o365Service.ModuleName = "AIPService"
#                if ($null -eq ($o365Service.ModuleInfo = Get-Module -Name $o365Service.ModuleName -ListAvailable -Verbose:$verboseFlag))  {
#                    Write-Warning -Message ("[PROCESS] Skipping {0}! {1} module is not present." -f $o365Service.Name, $o365Service.ModuleName)
#                }
#                else {
#                    $o365Service.ModuleVersion = (($o365Service.ModuleInfo).Version).ToString()
#                    $hashArgs = @{ Credential = $userCredential }
#                    switch ($true) {
#                        { $useMFA } { $hashArgs.Add('AccountId', $UserPrincipalName) ; $hashArgs.Remove('Credential') }
#                        { $TenantID } { $hashArgs.Add('TenantID', $TenantID) ; }
#                    }
#                    Import-Module -Name $o365Service.ModuleName -Verbose:$verboseFlag
#                    try { 
#                        Write-Verbose -Message ("[PROCESS] [TRY] Attepmting to connect to {0} ..." -f $o365Service.Name)
#                        $null = Connect-AIPService @hashArgs -ErrorAction Stop
#                        # if ($tID = Get-AzureADTenantDetail) { $o365Service += (" [{0}]" -f $tID.DisplayName) }
#                        Write-Verbose -Message ("[PROCESS] [TRY] {0} connected." -f $o365Service.Name)
#                        $o365Service.IsConnected = $true
#                        $o365Service.ConnectedAt = (Get-Date)
#                        $script:connectedO365Services += $o365Service
#                    }
#                    catch {
#                        Write-Warning -Message ("[PROCESS] [CATCH] {0} - Connection exception occured!" -f $o365Service)
#                        $PSCmdlet.ThrowTerminatingError($PSitem)
#                    }
#            }
            }
        }
    }
    end {
        if ($global:connectO365Services.Count -gt 0) {
            Write-Output "Available O365Services :",$global:connectO365Services
        }
        Write-Verbose -Message "[END] Ending $($MyInvocation.MyCommand)"
    }
}
Export-ModuleMember -Function Connect-O365Service
