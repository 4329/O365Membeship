$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path).Replace(".Tests.", ".")
# . "$here\$sut"


Describe "Connect to O365" -Tags "Connect-O365Admin" {
    Context "Existing User"{
        $testUser="kevin.sheck@buildrobot.org"
        $testDomain="buildrobot.org"
        $passfilename = $env:TEMP+"\O365Membership.txt"
        It "Connecting with prompting for password" {
            Remove-Item -Path $passfilename -ErrorAction SilentlyContinue
            Write-Host "Should be Prompted for password"
            .\Connect-O365Admin.ps1 -O365UserName $testUser
            Get-MsolDomain | Where-Object { $_.Name -eq $testDomain} | Should Not BeNullOrEmpty
            gsn | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"} | Should Not BeNullOrEmpty
            Test-Path -Path $passfilename  | Should Be $true
        }
        It "Connecting without prompting for password" {
            # Assumes passfile for user already exists
            Test-Path -Path $passfilename  | Should Be $true
            Write-Host "Should NOT be Prompted for password"
            .\Connect-O365Admin.ps1 -O365UserName $testUser
            Get-MsolDomain | Where-Object { $_.Name -eq $testDomain} | Should Not BeNullOrEmpty
            gsn | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"} | Should Not BeNullOrEmpty
            Test-Path -Path $passfilename | Should Be $true
        }
    }
}