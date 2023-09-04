$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

$groupsToDelete = $form.gridGroups

<#----- Exchange On-Premises: Start -----#>
# Connect to Exchange
try{
    $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername,$adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -ErrorAction Stop 
    $session = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber
    Write-Information "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" 
    
    $Log = @{
            Action            = "DeleteResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log 

} catch {
    Write-Error "Error connecting to Exchange using the URI [$exchangeConnectionUri]. Error: $($_.Exception.Message)"
    $Log = @{
            Action            = "DeleteResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Failed to connect to Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}

foreach($distributiongroup in $groupsToDelete){
    try{
        
        Remove-DistributionGroup -Identity $($distributiongroup.DN) -Confirm:$false
        Write-Information "Successfully deleted distributiongroup [$($distributiongroup.DisplayName)]" 
        
        $Log = @{
                Action            = "DeleteResource" # optional. ENUM (undefined = default) 
                System            = "Exchange On-Premise" # optional (free format text) 
                Message           = "Successfully deleted distributiongroup [$($distributiongroup.DisplayName)]" # required (free format text) 
                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $distributiongroup.DisplayName # optional (free format text) 
                TargetIdentifier  = $([string]$distributionGroup.UserPrincipalName) # optional (free format text) 
            }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log   
        
    }catch{
        Write-Error "Error deleting distributiongroup [$($distributiongroup.DisplayName)].  Error: $($_.Exception.Message)"
        $Log = @{
                Action            = "DeleteResource" # optional. ENUM (undefined = default) 
                System            = "Exchange On-Premise" # optional (free format text) 
                Message           = "Error deleting distributiongroup for [$($distributiongroup.DisplayName)]." # required (free format text) 
                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $distributiongroup.DisplayName # optional (free format text) 
                TargetIdentifier  = $([string]$distributionGroup.UserPrincipalName) # optional (free format text) 
            }
        #send result back  
        Write-Information -Tags "Audit" -MessageData $log 
    }
}

# Disconnect from Exchange
try{
    Remove-PsSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
    Write-Information "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]"     
    $Log = @{
            Action            = "DeleteResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
            IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log 
} catch {
    Write-Error "Error disconnecting from Exchange.  Error: $($_.Exception.Message)"
    $Log = @{
            Action            = "DeleteResource" # optional. ENUM (undefined = default) 
            System            = "Exchange On-Premise" # optional (free format text) 
            Message           = "Failed to disconnect from Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
            IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
            TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
            TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
        }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log    
    
}
<#----- Exchange On-Premises: End -----#>
