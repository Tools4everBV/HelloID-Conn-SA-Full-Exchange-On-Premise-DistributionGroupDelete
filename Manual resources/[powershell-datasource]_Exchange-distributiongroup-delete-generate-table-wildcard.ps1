try {
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $adminSecurePassword)
    $searchValue = ($dataSource.searchValue).trim()
    $searchQuery = "*$searchValue*"  

    $sessionOptionParams = @{
        SkipCACheck = $false
        SkipCNCheck = $false        
    }

    $sessionOption = New-PSSessionOption  @SessionOptionParams 

    $sessionParams = @{
        Authentication    = 'Default' 
        ConfigurationName = 'Microsoft.Exchange' 
        ConnectionUri     = $ExchangeConnectionUri 
        Credential        = $adminCredential        
        SessionOption     = $sessionOption       
    }

    $exchangeSession = New-PSSession @SessionParams

    Write-Information "Search query is '$searchQuery'" 
    
    $getDistributionGroupParams = @{        
        Filter = "alias -like '$searchQuery' -or name -like '$searchQuery' -or displayname -like '$searchQuery'"   
    }
   
    $invokecommandParams = @{
        Session      = $exchangeSession
        Scriptblock  = [scriptblock] { Param ($Params)Get-DistributionGroup @Params }
        ArgumentList = $getDistributionGroupParams
    }

    Write-Information "Successfully connected to Exchange '$ExchangeConnectionUri'"  
    
    $distributionGroups = Invoke-Command @invokeCommandParams   

    $distributionGroups = $distributionGroups | Sort-Object -Property DisplayName
        
    $resultDistributionGroups = [System.Collections.Generic.List[PSCustomObject]]::New()
    foreach ($group in $distributionGroups) {        
        $resultGroup = @{
            DisplayName       = $group.DisplayName
            UserPrincipalName = $group.PrimarySMTPAddress
            DN                = $group.DistinguishedName

        }
        $resultDistributionGroups.add($resultGroup)

    }
    $resultDistributionGroups
    
    Remove-PSSession($exchangeSession)
  
}
catch {
    Write-Error "Error connecting to Exchange using the URI '$exchangeConnectionUri', Message '$($_.Exception.Message)'"
}

