#Global Pramateters needed for Signature creation 
param
(
	[Parameter(Mandatory = $true,ValueFromPipeline = $True,ValueFromPipelinebyPropertyName = $True)] [string]$Sig1,
	[Parameter(Mandatory = $true,ValueFromPipeline = $True,ValueFromPipelinebyPropertyName = $True)] [string]$Sig2,
	[Parameter(Mandatory = $true,ValueFromPipeline = $True,ValueFromPipelinebyPropertyName = $True)] [string]$TemplatePath
)

function Write-log {
	[CmdletBinding()]

	param([Parameter(Position = 0)][ValidateNotNullOrEmpty()] [string]$Message,
		[Parameter(Position = 1)] [string]$Logfile = $log
	)

	Write-Output "$(Get-Date) $Message" | Out-File -FilePath $LogFile -Append
	Write-Verbose $message
}
function Get-FileEncoding {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $True,ValueFromPipelinebyPropertyName = $True)]
		[string]$Path
	)

	[byte[]]$byte = Get-Content -Encoding byte -ReadCount 4 -TotalCount 4 -Path $Path
	#Write-Host Bytes: $byte[0] $byte[1] $byte[2] $byte[3]

	# EF BB BF (UTF8)
	if ($byte[0] -eq 0xef -and $byte[1] -eq 0xbb -and $byte[2] -eq 0xbf)
	{ $Encoding = 'UTF8' }

	# FE FF  (UTF-16 Big-Endian)
	elseif ($byte[0] -eq 0xfe -and $byte[1] -eq 0xff)
	{ $Encoding = 'Unicode UTF-16 Big-Endian' }

	# FF FE  (UTF-16 Little-Endian)
	elseif ($byte[0] -eq 0xff -and $byte[1] -eq 0xfe)
	{ $Encoding = 'Unicode UTF-16 Little-Endian' }

	# 00 00 FE FF (UTF32 Big-Endian)
	elseif ($byte[0] -eq 0 -and $byte[1] -eq 0 -and $byte[2] -eq 0xfe -and $byte[3] -eq 0xff)
	{ $Encoding = 'UTF32 Big-Endian' }

	# FE FF 00 00 (UTF32 Little-Endian)
	elseif ($byte[0] -eq 0xfe -and $byte[1] -eq 0xff -and $byte[2] -eq 0 -and $byte[3] -eq 0)
	{ $Encoding = 'UTF32 Little-Endian' }

	# 2B 2F 76 (38 | 38 | 2B | 2F)
	elseif ($byte[0] -eq 0x2b -and $byte[1] -eq 0x2f -and $byte[2] -eq 0x76 -and ($byte[3] -eq 0x38 -or $byte[3] -eq 0x39 -or $byte[3] -eq 0x2b -or $byte[3] -eq 0x2f))
	{ $Encoding = 'UTF7' }

	# F7 64 4C (UTF-1)
	elseif ($byte[0] -eq 0xf7 -and $byte[1] -eq 0x64 -and $byte[2] -eq 0x4c)
	{ $Encoding = 'UTF-1' }

	# DD 73 66 73 (UTF-EBCDIC)
	elseif ($byte[0] -eq 0xdd -and $byte[1] -eq 0x73 -and $byte[2] -eq 0x66 -and $byte[3] -eq 0x73)
	{ $Encoding = 'UTF-EBCDIC' }

	# 0E FE FF (SCSU)
	elseif ($byte[0] -eq 0x0e -and $byte[1] -eq 0xfe -and $byte[2] -eq 0xff)
	{ $Encoding = 'SCSU' }

	# FB EE 28  (BOCU-1)
	elseif ($byte[0] -eq 0xfb -and $byte[1] -eq 0xee -and $byte[2] -eq 0x28)
	{ $Encoding = 'BOCU-1' }

	# 84 31 95 33 (GB-18030)
	elseif ($byte[0] -eq 0x84 -and $byte[1] -eq 0x31 -and $byte[2] -eq 0x95 -and $byte[3] -eq 0x33)
	{ $Encoding = 'GB-18030' }

	else
	{ $Encoding = 'ASCII' }
	return $Encoding
}
function Set-OutlookSig {
	################## 
	#Script Variables# 
	################## 
	$UserName = $env:username
	$LocalAppData = $env:LOCALAPPDATA
	$Date = Get-Date
	$log = "$LocalAppData\OutlookSig_$UserName.log"
    $UseThumbnailPhoto = $false

	Write-log "----------------------------------------------------------------------------------------------------------------------------------------"
	Write-log "----------------------------------------------------------------------------------------------------------------------------------------"
	Write-log "#################################################################################################"
	Write-log "#######################################-Start of Signature Script-#######################################"
	Write-log "#################################################################################################"
	############################################################################################## 
	#Get Office Version 
	$OfficeVersion = (Get-ItemProperty -Path Registry::HKEY_CLASSES_ROOT\Outlook.Application\CurVer). "(Default)".split('.')[2] + ".0"
	#Get Name of Outlook Signature Folder 
	$OutlookSigDirName = (Get-ItemProperty -Path Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\$OfficeVersion\Common\General). "Signatures"
	#Build Path for Outlook Signature 
	$OutlookSigLocalPath = "${env:appdata}\Microsoft\$OutlookSigDirName"
    #Build User Picture Path
    $ADUserPicturePath = $OutlookSigLocalPath+"\"+"$UserName"+"_Thumbnail.jpg"

	############################################################################################## 

	#Cleanup old Signatures 
	if (!(Test-Path -Path $log)) {
		Remove-Item $OutlookSigLocalPath\*.* -Force
		Write-log "Signatures Folder cleaned" }

	#Remove Logfile if its to big 
	$logsize = if ((Get-Item $log).Length -gt 5mb) {
		Remove-Item $log -Force -Confirm:$false
		Write-log "Old Logfile Removed" }
	############################################################################################## 

	#Write-log "##################################################################################" 
	Write-log "Office Version detected: $OfficeVersion"
	Write-log "Outlook Signature Foldername detected: $OutlookSigDirName"
	Write-log "Build Outlook Signature Path: $OutlookSigLocalPath"

	############################################################################################## 

	#Build Signature Content 
	$Sig1Content = "$TemplatePath" + "$Sig1" + ".htm"
	$Sig2Content = "$TemplatePath" + "$Sig2" + ".htm"

	Write-log "##################################################################################"
	Write-log "Signature Script Version: $Version"
	Write-log "Create Signature started on: $Date"
	Write-log "Create Signature: $Sig1"
	Write-log "Create Signature: $Sig2"
	Write-log "##################################################################################"


	#AD Searcher Part 
	$Ldap = "dc=domain,dc=local"
	$Filter = "(&(objectCategory=User)(samAccountName=$UserName))"
	$Searcher = [adsisearcher]$Filter
	$searcher.SearchRoot = "LDAP://$Ldap"
	$ADUserPath = $Searcher.FindOne()
	$ADUser = $ADUserPath.GetDirectoryEntry()

	################################### 
	#Get required AD User Informations# 
	################################### 

	#Build AD Displayname with Surname and Givenname 
	$ADDisplayName = $ADUser.givenName + $ADUser.sn

	#Get AD title if it is Present 
	$ADTitle = $ADUser.Title
	if ($ADTitle) { $ADTitleValue = $ADTitle }

	#Get AD Department if it is present and split out '^' to a new Value 
	$ADDepartment = $ADUser.department
	if ($ADDepartment) { $ADDepartmentValue = $ADDepartment }
	if ($ADDepartment -like '*^*') { $ADDepartmentValue = $ADDepartmentValue.split('^')[-2]; $SecondDepartment = $ADDepartment.split('^')[-1] }

	#Get AD Mobile Number if it is Present 
	$ADMobile = $ADUser.Mobile
	if ($ADMobile) { $ADMobileValue = "Mobile: " + "$ADMobile" }

	#Get AD FAX Number if it is Present 
	$ADFax = $ADUser.facsimileTelephoneNumber
	if ($ADFax) { $ADFaxValue = "Fax: " + "$ADFax" }

	#Get AD Phone Number if it is Present 
	$AdPhone = $ADUser.telephoneNumber
	if ($AdPhone) { $AdPhoneValue = "Tel.: " + "$AdPhone" }

	#Get AD E-Mail 
	$ADemail = $ADUser.mail

	#Get AD Description 
	$ADDescription = $ADUser.description

    #Get AD Thumbnail Photo
    $ADThumbnail = $ADUser.thumbnailPhoto

	#Write AD Search output to Logfile 
	Write-log "-------------------------------------------------------------------------------------------------------"
	Write-log "Read User Informations from AD"
	Write-log "-------------------------------------------------------------------------------------------------------"
	Write-log "User Name: $ADDisplayName"
	Write-log "User Title: $ADTitle"
	Write-log "Department: $ADDepartment"
	Write-log "Mobile Number: $ADMobile"
	Write-log "Fax Number: $ADFax"
	Write-log "Telephone Number: $AdPhone"
	Write-log "Mail: $ADemail"
	Write-log "Description: $ADDescription"
    
	############################ 
	#Create the Signature files# 
	############################ 

	#Write Signature Path to Logfile 
	Write-log "-------------------------------------------------------------------------------------------------------"
	Write-log "Signatures Folder:  $OutlookSigLocalPath"
	Write-log "-------------------------------------------------------------------------------------------------------"


	#Get Item Encoding
	$EncodingSig1 = Get-FileEncoding $Sig1Content
	$EncodingSig2 = Get-FileEncoding $Sig2Content

	Write-log "Get Template Encodings..."
	Write-log "File Encoding - $Sig1 : $EncodingSig1"
	Write-log "File Encoding - $Sig2 : $EncodingSig2"

	#Create the Signature Folder if its not present
	if (!(Test-Path -Path $OutlookSigLocalPath)) { mkdir $OutlookSigLocalPath }

    #Export User AD Picture to Signature Folder
    if($UseThumbnailPhoto -eq $True){
    $ADThumbnail | Set-Content $ADUserPicturePath -Encoding Byte
    Write-log "Thumbnail Picture Exported..."
    }

	#Get Signature Content and fill with Values from AD Search

	#############
	#Signature 1#
	#############
	#Get Signature Content and fill with Values from AD Search
	if ($EncodingSig1 -eq 'ASCII') {
		Invoke-Expression ('$Sig_1 = @"' + "`n" + (Get-Content -Path "$Sig1Content" | ForEach-Object { $_ + "`n" }) + "`n" + '"@')
		Write-log "Get Signature Content..."
		#Export Signature with Values to Soignature Path
		$Sig_1 | Out-File "$OutlookSigLocalPath\$Sig1.htm" -Force -Confirm:$false
		Write-log "Out finished Signature template File..."
	}
	else {
		Invoke-Expression ('$Sig_1 = @"' + "`n" + (Get-Content -Encoding $EncodingSig1 -Path "$Sig1Content" | ForEach-Object { $_ + "`n" }) + "`n" + '"@')
		Write-log "Get Signature Content..."
		#Export Signature with Values to Soignature Path
		$Sig_1 | Out-File -Encoding $EncodingSig1 "$OutlookSigLocalPath\$Sig1.htm" -Force -Confirm:$false
		Write-log "Out finished Signature template File..."
	}

	#############
	#Signature 2#
	#############

	if ($EncodingSig2 -eq 'ASCII') {
		Invoke-Expression ('$Sig_2 = @"' + "`n" + (Get-Content -Path "$Sig2Content" | ForEach-Object { $_ + "`n" }) + "`n" + '"@')
		Write-log "Get Signature Content..."
		#Export Signature with Values to Soignature Path
		$Sig_2 | Out-File "$OutlookSigLocalPath\$Sig2.htm" -Force -Confirm:$false
		Write-log "Out finished Signature template File..."
	}
	else {
		Invoke-Expression ('$Sig_2 = @"' + "`n" + (Get-Content -Encoding $EncodingSig2 -Path "$Sig2Content" | ForEach-Object { $_ + "`n" }) + "`n" + '"@')
		Write-log "Get Signature Content..."
		#Export Signature with Values to Soignature Path	
		$Sig_2 | Out-File -Encoding $EncodingSig2 "$OutlookSigLocalPath\$Sig2.htm" -Force -Confirm:$false
		Write-log "Out finished Signature template File..."
	}
	#Write to Logfile which Signature was created 
	Write-log "Signature: $Sig1.htm created"
	Write-log "Signature: $Sig2.htm created"
	Write-log "##################################################################################################"
	Write-log "########################################-End of Signature Script########################################"
	Write-log "##################################################################################################"
}

########################################## 
###############Script Version############# 
########################################## 

$Version = "1.5"

#Run Script with Parameters 
Set-OutlookSig -SigName $Sig1 -SigReplyName $Sig2 -TemplatePath $TemplatePath
