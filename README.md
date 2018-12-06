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
 
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"> 
<html xmlns="http://www.w3.org/1999/xhtml"> 
<head> 
<meta http-equiv="content-type" content="text/html; charset=Windows-1250" /> 
</head> 
<body> 
<div style="margin: 0px; padding: 0px; font-size: 9pt; font-family: Arial; color: black;"> 
    <p style="margin: 0px; margin-bottom: 12px;">Greetings,</p> 
    <p style="margin: 0px; text-transform: uppercase; font-weight: bold;">$ADDisplayName</p> 
    <p style="margin: 0px; margin-bottom: -8px;">$ADTitleValue</p> 
    <p style="margin: 0px; margin-bottom: -8px;">$ADDepartmentValue</p> 
    <p style="margin: 0px; margin-bottom: -8px;">$SecondDepartment</p> 
    <p style="margin: 0px;">_____________________________</p> 
    <p style="margin: 0px;">$AdPhoneValue</p> 
    <p style="margin: 0px;">$ADFaxValue</p> 
    <p style="margin: 0px;"><a href="mailto:$ADemail">$ADemail</a></p> 
    </p> 
</div> 
</body> 
</html>

Example on How to run the Script:
SignatureScript.ps1 -TemplatePath '\\server0001\Templates\' -Sig1 'Signame' -Sig2 'Sigreplyname'
Evreything in Script is build in Windows, so no Modules or wrapper needed.
Have Fun :)
