Param
    (

        [Parameter(Mandatory=$false,
                   Position=0)]$O365UserName
    )
Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Entering")


$error.Clear()
import-Module -Name MSOnline 
if($error.Count -gt 0){
    Write-Error -message "Load of MSOnline Failed - Make sure Microsoft Online Services Sign-in Assistant is installed AND running x64 PowerShell Instance" -RecommendedAction "Install MIcrosoft Online Servides Sign-in Assistant"
    exit 1
}

#Connect for MSOL
$passfilename = $env:TEMP+"\O365Membership.txt"
if(Test-Path  $filename){
    Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Using Existing Credentials")
    $password= Get-Content $filename | ConvertTo-SecureString
}
else{
    Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"NOT using existing credentials")
    $password=$null
}


#Clear errors to try to determine if problems occured getting credentials, like hitting cancel when prompted
$error.Clear()
$BRCredential=$null

if($O365UserName -ne $null){
	Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("O365UserName provided is {0}" -f $O365UserName))
    $passfilename = $env:TEMP+"\O365Membership.txt"
    if(Test-Path  $passfilename){
        Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Using Existing Password from {0}" -f $passfilename ))
        Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),(" Generating Credentials for User {0}" -f $O365UserName))
        $BRCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $O365UserName,(Get-Content $passfilename | ConvertTo-SecureString)
    }
    else{
        Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Prompting for Credentials for User {0}" -f $O365UserName))
        $BRCredential= Get-Credential -Message "Sign In as O365 Admin" -UserName $O365UserName
        Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Saving Password {0}" -f $passfilename))
        $BRCredential.Password | ConvertFrom-SecureString | Out-File $passfilename
    }
}
else{
	Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"No O365UserName provided so prompt for User and Password")
    $BRCredential= Get-Credential -Message "Sign In as O365 Admin"
}
if(($error.count -gt 0) -or ($BRCredential -eq $null)) {
	throw ($MyInvocation.InvocationName+" - "+ "Unable to Get-Credential - Can't continue")
	exit 1
}


Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Connecting to MSOnline")
try {
    connect-msolservice -Credential $BRCredential
}
catch{
    Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Error {0} connecting to MSOnline" -f $error[0].ErrorDetails))
    exit 2
}

Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Connecting to Exchange Online")
try{
    if ((gsn | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"}) -eq $null){
        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $BRCredential -Authentication Basic -AllowRedirection
        Import-PSSession -Session $ExchangeSession -AllowClobber
    }
    else{
        Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Found exising Exchange Online Session")
    }
}
catch{
    Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Error {0} connecting to Exchange Session" -f $error[0].ErrorDetails))
    exit 2
    
}



