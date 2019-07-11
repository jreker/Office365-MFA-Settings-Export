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
$CSVPath = Read-Host "Where do you want to store the CSV file? (eg. C:\temp\export.csv)"

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

function getResult() {
    try{
        $result = New-Object Collections.Generic.List[PSObject]
        #Get All MSOL Users
        write-host "Start to export user settings...please wait..."
        foreach($user in Get-MsolUser -All -ErrorAction Stop) {
            
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
#export results to csv
getResult | Export-Csv -Path $CSVPath -NoClobber -Delimiter ';'

