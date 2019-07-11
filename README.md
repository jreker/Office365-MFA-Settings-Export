# Office 365 MFA Settings Exporter
This powershell script exports the mfa settings of all office 365 users of your tenant.
The export format is .csv with the delemiter ';'.

## Prerequisites
- Installed powershell module "MSOnline".
    - If you dont have it installed just install it by ```Install-Module MSOnline``` in admin powershell session.
- Admincredentials for the tentant where you want to export the settings.
- Excel or another software to open/sort and filter the csv file.

## Exported Attributes
The script will export the following attributes:

- FirstName :: The first name of the user
- LastName :: The last name of the user
- UserPrincipalName :: The upn of the user
- MFA_State :: The MFA state of the user (deactivated;enabled;enforced)
- MFA_ActivationDate :: The date when MFA was activated for the user
- DefaultStrongAuthenticationMethodType :: The default selected authentication option/method
- PhoneAppAuthenticationTypes :: The selected authentication type of the phone (OTP/ App-Notification)
- PhoneDeviceName :: The name of the users phone 
- PhoneAppVersion :: The version of the users authenticator app
- PhoneAppNotificationType :: The notification type of the phone app
- PhoneNumber :: The phone number of the users configured mfa setup
- AlternativePhoneNumber :: The alternative phone number of the users configured mfa setup

*When a attribute is not available or not configured by the user it will be blank in the list!*

Feel free to contribute!
