##########################################################
# This is a script that exports all Office 365 Users and 
# the configuration of the MFA in O365.
#
# Author: jreker
# Date: 11.07.2019
# 
###########################################################


Import-Module MSOnline
##### Global Variables
$global:credential = $null

##### USER INPUTS
write-host "--Office 365 MFA-Settings exporter --"


#Function that checks if you are already connected to Office 365 
function connectToMSO() {
    try
    { 
        #check if already logged in
        $temp = Get-MsolAccountSku -ErrorAction Stop 
    } catch {
        $global:credential = Get-Credential -Message "Please type in your office 365 admin credentials."
        Connect-MsolService -Credential $global:credential
    } 
}

function getResult($userUPN) {
    try{
        $result = New-Object Collections.Generic.List[PSObject]
        #Get All MSOL Users
        write-host "Start to export user settings...please wait..."
        $users = ""
        if($null -eq $userUPN) {
            $users = Get-MsolUser -All -ErrorAction Stop
        } else {
            $users = Get-MsolUser -UserPrincipalName $userUPN
        }
        
        foreach($user in $users) {
            
            $MFA_STATE = $null
            if($null -eq $user.StrongAuthenticationRequirements.State){ 
                $MFA_STATE = "deactivated"
            } else {
                $MFA_STATE = $user.StrongAuthenticationRequirements.State
            }

            $usrObj = [PSCustomObject] @{
                FirstName = $user.FirstName
                LastName = $user.LastName
                UserPrincipalName = $user.UserPrincipalName
                MFA_State = $MFA_STATE
                MFA_ActivationDate = $user.StrongAuthenticationRequirements.RememberDevicesNotIssuedBefore
                DefaultStrongAuthenticationMethodType = ($user.StrongAuthenticationMethods | where-object { $_.IsDefault -eq $True }).MethodType
                PhoneAppAuthenticationTypes = $user.StrongAuthenticationPhoneAppDetails.AuthenticationType
                PhoneDeviceName = $user.StrongAuthenticationPhoneAppDetails.DeviceName
                PhoneAppVersion = $user.StrongAuthenticationPhoneAppDetails.PhoneAppVersion
                PhoneAppNotificationType = $user.StrongAuthenticationPhoneAppDetails.NotificationType
                PhoneNumber = $user.StrongAuthenticationUserDetails.PhoneNumber
                AlternativePhoneNumber = $user.StrongAuthenticationUserDetails.AlternativePhoneNumber                
            }
            $result.Add($usrObj)
        }
        
    } catch {
        write-host "Error running Get-MsolUser. Try again."
    }
    write-host "done"
    return $result
}

#connect to microsoft with office 365 credentials
connectToMSO


function Show-Menu
{
     param (
           [string]$Title = 'Office 365 # MFA Export'
     )
     Clear-Host
     Write-Host "================ $Title ================"
     Write-Host "1: Press '1' to export MFA settings of all users to csv."
     Write-Host "2: Press '2' to show settings of specific user." 
     Write-Host "Q: Press 'Q' to quit."
}



do
{
     Show-Menu
     $input = Read-Host "Please make a selection"
     switch ($input)
     {
           '1' {
                Clear-Host
                Write-Host "==> 1: export MFA settings of all users."
                $CSVPath = Read-Host "Where do you want to store the CSV file? (eg. C:\temp\export.csv)"
                getResult | Export-Csv -Path $CSVPath -NoClobber -Delimiter ';'
           } '2'  {
                Clear-Host
                Write-Host "==> 2:show M FA settings of specific user"
                $userUPN = read-host "UPN of user"
                getResult $userUPN
           } 'q' {
            return
       }
 }
 pause
}
until ($input -eq 'q')




#export results to csv


