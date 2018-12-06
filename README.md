# Outlook-Signature
This is a Script that will Search AD for Signature relevant Attributs and put them into Variables.

Then inserts those into a HTM Template that will be copied to Outlook Signature Folder (Multilanguage Support).

    ADDisplayName
    ADTitle
    ADDepartment
    ADMobile
    ADFax
    AdPhone
    ADemail
    ADDescription 

You need 3 Things:

Signame1, Signame2 (reply) and TemplatePath - where the HTM Template Files is Stored.

 

Version 1.5

    User ThumbnailPicture Export   to Signature Path , set $UseThumbnailPhoto on Line 85 to '$true' 

 

Version 1.4

    Added a new function "Get-FileEncoding" and Updated the Script Code.
    The Get-Content and Out-File have now Encoding Support. 

 

Example for Sig Template File:

![alt text](https://raw.githubusercontent.com/FSCorrupt/Outlook-Signature/master/html.png)

Example on How to run the Script:

SignatureScript.ps1 -TemplatePath '\\server0001\Templates\' -Sig1 'Signame' -Sig2 'Sigreplyname'

Evreything in Script is build in Windows, so no Modules or wrapper needed.
Have Fun :)
