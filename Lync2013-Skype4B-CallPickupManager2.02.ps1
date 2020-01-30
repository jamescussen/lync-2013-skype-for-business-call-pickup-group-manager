########################################################################
# Name: Lync 2013 / Skype4B Call Pickup Group Manager 
# Version: v2.02 (26/1/2019)
# Date: 11/10/2013
# Created By: James Cussen
# Web Site: http://www.myskypelab.com
# 
# Notes: This is a PowerShell tool. To run the tool, open it from the PowerShell command line on a Lync / Skype for Business server.
#		 For more information on the requirements for setting up and using this tool please visit http://www.myskypelab.com.
#
# Copyright: Copyright (c) 2019, James Cussen (www.myskypelab.com) All rights reserved.
# Licence: 	Redistribution and use of script, source and binary forms, with or without modification, are permitted provided that the following conditions are met:
#				1) Redistributions of script code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				2) Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				3) Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
#				4) This license does not include any resale or commercial use of this software.
#				5) Any portion of this software may not be reproduced, duplicated, copied, sold, resold, or otherwise exploited for any commercial purpose without express written consent of James Cussen.
#			THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; LOSS OF GOODWILL OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#
# Prerequisites:
#	- Lync 2013: Requires SEFAUtil installed on the system and SQL Dynamic Ports opened in Windows Firewall on all Front End servers.
#	- Skype for Business (CU1+): Requires SQL Dynamic Ports opened in Windows Firewall on all Front End servers.
#	- The SQL ports on Front End servers firewall can be opened automatically with the OpenSQLPortsForCallPickupManager1.00.ps1 supplied with this tool.
#	- Get more information here: http://www.myskypelab.com/2013/10/lync-2013-call-pickup-group-manager.html
#	
#
# Release Notes:
# 1.00 Initial Release.
#
# 1.01 Feature update
#		- Pre-Req check will now look under the default reskit location on all available drives (not just C:)
#		- If SEFAUTIL gives no response (due to an unknown error in SEFAUTIL) the tool will display an error to the user.
#		- Added the Import-Module Lync command in case you run the script from regular Powershell or use the Right Click Run using Powershell method to start the script.
#
# 1.02 Maintenance Update
#		- Added the undocumented "/verbose" flag to the SEFAUtil calls to help with debugging SEFAUtil issues. See post: http://www.myskypelab.com/2014/04/sefautil-and-lync-2013-call-pickup.html
#
# 1.03 Common Area Phone Update
#		- This version has been updated to handle Common Area Phones. Some people reported errors being displayed by the tool when they had manually set (with SEFAUTIL) Group Call Pickup against Common Area Phones (ie. against the SIP URI of the Common Area Device, eg: sip:fbcb642b-f5bc-477a-a053-373aef4c00f8@domain.com). As of this version Common Area Phones will be included in the user list, and you can add and remove them from Call Pickup Groups.
#		- User listboxes are now slightly wider to deal with the long SIP addresses of Common Area Phones.
#		- When the tool loads it will display in the PowerShell window the SIP Address and Display Name of all common area devices so you can match the (GUID looking) SIP address in the tool to the display name of the device.
#
# 1.04 Scalability Update
#		- Now supports window resizing.
#		- Added Filter on Lync users listbox to cater for deployments that have lots of users.
#		- Script is now signed.
#	
# 1.05 Enhancements
#		- Now checks that Group number matches a range that exists in one of the Call Pickup Orbits before allowing it to be added to the group list.
#		- Up and Down keys in Orbit listbox now update orbit details properly.
#		- You can now specify an alternate location of SEFAUTIL.exe in the command line. (Example: .\Lync2013CallPickupManager.1.05 "D:\folder\SEFAUTIL.exe")
#		- Now checks the Skype for Business RESKIT location.
#		- Put a dividing line between the Orbit creation section and the Group configuration section to try and indicate a divide between the two areas.
#		- The Group list box label now displays the group name being listed up so the user understands better what group the user list is associated with.
#
# 2.00 Major Update
#		- When using Skype for Business the new Call Pickup Group Powershell commands (Available in CU1+) for user settings are detected and used instead of SEFAUTIL. This means you don't have to worry about any SEFAUTIL configuration anymore in Skype for Business CU1 or higher!
#		- Unfortunately the Skype for Business Powershell commands have proved to be too slow for the discovery of all users Call Pickup Settings (because "Get-CsUser | Get-CsGroupPickupUserOrbit" has to iterate through all users taking about 2+ seconds per user and as a result takes ages for large sites). So I have retained and made major improvements to the direct SQL discovery method from version 1.0 for both Lync 2013 and Skype for Business.
#		- Changed the way the Groups list box works. It now will be automatically filled with all the available groups from the Orbit ranges assigned in the system. If there is a user in a Group then it will be highlighted in Green text with Yellow background to help you find users without looking in empty groups.
#		- When Orbits are added all the available groups (eg. from Range Start to Range End) will automatically be added to the available Groups list. Note: when you remove ranges that contain groups with users in them the group will no longer be accessible in the interface, however, the Pickup Group will still function. So make sure you remove users from groups before you delete the Orbit range.
#		- Added a refresh button so the data can now be updated from the system whenever you want.
#		- Improved the speed of looking through groups by re-architecting the group coding.
#		- After a user is added or removed from a group the tool does not rescan all user's data again. Whilst fully rescanning always ensures the data being displayed is up to date, it's also much slower. This version is optimised for speed :)
#		- Added pretty looking UP and DOWN arrow icons on add and remove buttons to try and clarify operation.
#		- Added the "Find Selected User" button to find which group a user is assigned to.
#		- Added help tool tips on buttons.
#
# 2.01 Enhancements
#		- Added an export to CSV data to the app.
#		- Fixed an issue with highlighting groups with users in them.
#
# 2.02 Enhancements
#		- Updated for Skype for Business 2019
#
########################################################################

param (
    [string]$SEFAUtilLocation = ""
)


$theVersion = $PSVersionTable.PSVersion
$MajorVersion = $theVersion.Major

Write-Host ""
Write-Host "--------------------------------------------------------------"
Write-Host "PowerShell Version Check..." -foreground "yellow"
if($MajorVersion -eq  "1")
{
	Write-Host "This machine only has Version 1 PowerShell installed.  This version of PowerShell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "2")
{
	Write-Host "This machine has Version 2 PowerShell installed. This version of PowerShell is not supported." -foreground "red"
}
<#
elseif($MajorVersion -eq  "3")
{
	Write-Host "This machine has version 3 PowerShell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "4")
{
	Write-Host "This machine has version 4 PowerShell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "5")
{
	Write-Host "This machine has version 4 PowerShell installed. CHECK PASSED!" -foreground "green"
}
else
{
	Write-Host "This machine has version $MajorVersion PowerShell installed. Unknown level of support for this version." -foreground "yellow"
}
#>
Write-Host "--------------------------------------------------------------"
Write-Host ""


Function Get-MyModule 
{ 
Param([string]$name) 
	
	if(-not(Get-Module -name $name)) 
	{ 
		if(Get-Module -ListAvailable | Where-Object { $_.name -eq $name }) 
		{ 
			Import-Module -Name $name 
			return $true 
		} #end if module available then import 
		else 
		{ 
			return $false 
		} #module not available 
	} # end if not module 
	else 
	{ 
		return $true 
	} #module already loaded 
} #end function get-MyModule 


$Script:LyncModuleAvailable = $false
$Script:SkypeModuleAvailable = $false

Write-Host "--------------------------------------------------------------"
Write-Host "Importing Modules..." -foreground "yellow"
#Import Lync Module
if(Get-MyModule "Lync")
{
	Invoke-Expression "Import-Module Lync"
	Write-Host "Imported Lync Module..." -foreground "green"
	$Script:LyncModuleAvailable = $true
}
else
{
	Write-Host "Unable to import Lync Module..." -foreground "yellow"
}
#Import SkypeforBusiness Module
if(Get-MyModule "SkypeforBusiness")
{
	Invoke-Expression "Import-Module SkypeforBusiness"
	Write-Host "Imported SkypeforBusiness Module..." -foreground "green"
	$Script:SkypeModuleAvailable = $true
}
else
{
	Write-Host "Unable to import SkypeforBusiness Module... (Expected on a Lync 2013 system)" -foreground "yellow"
}
Write-Host "--------------------------------------------------------------"



$Script:SkypeForBusinessAvailable = $false

# Check for Skype for Business Commands
$command = "Set-CsGroupPickupUserOrbit"
if(Get-Command $command -errorAction SilentlyContinue)
{
	Write-Host
	Write-Host "--------------------------------------------------------------"
	Write-Host "Set-CsGroupPickupUserOrbit is available. This is Skype for Business CU1+" -foreground "green"
	Write-Host "Note: Common Area Phones are not supported by the Skype for Business PowerShell commands in CU1" -foreground "yellow"
	Write-Host "--------------------------------------------------------------"
	Write-Host
	$Script:SkypeForBusinessAvailable = $true
}
else
{
	Write-Host
	Write-Host "--------------------------------------------------------------"
	Write-Host "INFO: Set-CsGroupPickupUserOrbit command is not available. Tool will fall back to using SEFAUTIL commands if available." -foreground "yellow"
	Write-Host "--------------------------------------------------------------"
	Write-Host
	$Script:SkypeForBusinessAvailable = $false
}

if(!$Script:SkypeForBusinessAvailable)
{
	# 2013 Prerequisites Check ==========================================================
	Write-Host ""
	Write-Host "INFO: Checking Lync 2013 Prerequisites..." -Foreground "yellow"

	if($SEFAUtilLocation -ne $null -and $SEFAUtilLocation -ne "")
	{
		if(!($SEFAUTILPath2013 -match ".exe"))
		{
			write-host ""
			write-host "ERROR: When supplying a path in the command line, the path must contain the full SEFAUTIL filename. Example: .\Lync2013CallPickupManager.1.05 `"D:\folder\SEFAUTIL.exe`"" -foreground "red"
			exit
		}
		if (!(Test-Path $SEFAUTILPath2013))
		{
			write-host ""
			write-host "ERROR: SEFAUTIL not found in location: $SEFAUTILPath2013" -foreground "red"
			exit
		}
		
		$SEFAUTILPath2013 = $SEFAUtilLocation
	}
	else
	{
		#Location of SEFAUTIL. This is the Lync 2013 standard location. If yours is different to this, then change this variable.
		$SEFAUTILPath2013 = "c:\Program Files\Microsoft Lync Server 2013\Reskit\SEFAUtil.exe"
		
		if (Test-Path $SEFAUTILPath2013)
		{
			write-host ""
			write-host "Checking Prerequisites: Found SEFAUTIL in location: $SEFAUTILPath2013" -foreground "green"
		}
		else
		{
			write-host "INFO: SEFAUTIL not found. Checking other drives..." -foreground "yellow"		
			$AllDrives = Get-PSDrive -PSProvider 'FileSystem'
			$FoundDrive = $false
			foreach($Drive in $AllDrives)
			{
				[string]$DriveName = $Drive.Root
				$TestDrive = "${DriveName}Program Files\Microsoft Lync Server 2013\Reskit\SEFAUtil.exe"
				write-host "INFO: Checking: $TestDrive" -foreground "yellow"
				
				if (Test-Path $TestDrive)
				{
					write-host "Found SEFAUTIL Drive. Using: $TestDrive" -foreground "green"
					$SEFAUTILPath2013 = $TestDrive
					$FoundDrive = $true
					break
				}
			}
			foreach($Drive in $AllDrives) #Override with newer Skype4B Snooper if available
			{
				
				[string]$DriveName = $Drive.Root
				$TestDrive = "${DriveName}Program Files\Skype for Business Server 2015\Reskit\SEFAUtil.exe"
				Write-Host "INFO: Checking: $TestDrive" -foreground "yellow"
				
				if (Test-Path $TestDrive)
				{
					write-host "Found SEFAUTIL Drive. Using: $TestDrive" -foreground "green"
					$SEFAUTILPath2013 = $TestDrive
					$FoundDrive = $true
					break
				}
			}
			if(!$FoundDrive)
			{
				Write-Host "ERROR: Could not find a drive with SEFAUTIL installed on it. Please install Lync 2013 reskit tools or Skype4B SEFAUTIL.exe file." -foreground "red" 
				Write-Host ""
				Write-Host "INFO: If you have installed SEFAUTIL in a non standard location you can specify the path from the command line. Example: .\Lync2013CallPickupManager.1.05 `"D:\folder\SEFAUTIL.exe`"" -foreground "red" 
				exit
			}
		}
	}


	#Check if SEFAUTIL is programmed as a trusted application
	$appIds = Get-CsTrustedApplication | select-object ApplicationId

	$foundAppId = $false
	foreach($appId in $appIds)
	{
		if($appId.ApplicationId -imatch "urn:application:sefautil")
		{
			$foundAppId = $true
		}
	}
	if($foundAppId -eq $true)
	{
		write-host ""
		write-host "Checking Prerequisites: SEFAUTIL is configured as a trusted application." -foreground "green"
	}
	else
	{
		write-host ""
		write-host "ERROR: SEFAUTIL does not appear to be installed as a Trusted Application. Please configure SEFAUTIL as a trusted application before using Call Pickup Manager." -foreground "red"
	}

}
else
{
	Write-Host "INFO: Skipped SEFAUTIL checks because Skype for Business has been detected." -foreground "yellow"
}



$script:groups = @()
$script:pools = @()
$script:computers = @()


#Select only single computer from a pool or single computer from the pool. (Paired Pools are still added as separtate machines)
Get-CsPool | where-object {$_.Services -like "Registrar*"} | select-object Computers | ForEach-Object {$computers +=  $_.Computers[0]}


write-host ""
foreach($computer in $computers)
{
	Write-Host "Discovered Computer: $computer" -Foreground "yellow"
}
Write-Host ""


# Set up the form  ============================================================
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Lync 2013 / Skype4B Call Pickup Manager v2.02"
$objForm.Size = New-Object System.Drawing.Size(625,560) 
$objForm.MinimumSize = New-Object System.Drawing.Size(625,560)
$objForm.StartPosition = "CenterScreen"
[byte[]]$WindowIcon = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 32, 0, 0, 0, 32, 8, 6, 0, 0, 0, 115, 122, 122, 244, 0, 0, 0, 6, 98, 75, 71, 68, 0, 255, 0, 255, 0, 255, 160, 189, 167, 147, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 0, 7, 116, 73, 77, 69, 7, 225, 7, 26, 1, 36, 51, 211, 178, 227, 235, 0, 0, 5, 235, 73, 68, 65, 84, 88, 195, 197, 151, 91, 108, 92, 213, 21, 134, 191, 189, 207, 57, 115, 159, 216, 78, 176, 27, 98, 72, 226, 88, 110, 66, 66, 34, 185, 161, 168, 193, 73, 21, 17, 2, 2, 139, 75, 164, 182, 106, 145, 170, 190, 84, 74, 104, 65, 16, 144, 218, 138, 138, 244, 173, 69, 106, 101, 42, 129, 42, 149, 170, 162, 15, 168, 168, 151, 7, 4, 22, 180, 1, 41, 92, 172, 52, 196, 68, 105, 130, 19, 138, 98, 76, 154, 27, 174, 227, 248, 58, 247, 57, 103, 175, 62, 236, 241, 177, 199, 246, 140, 67, 26, 169, 251, 237, 236, 61, 179, 215, 191, 214, 191, 214, 191, 214, 86, 188, 62, 37, 252, 31, 151, 174, 123, 42, 224, 42, 72, 56, 138, 152, 99, 191, 175, 247, 114, 107, 29, 172, 75, 106, 94, 254, 74, 156, 109, 13, 58, 180, 155, 53, 240, 216, 64, 129, 63, 156, 43, 95, 55, 0, 106, 62, 5, 158, 134, 83, 59, 147, 116, 36, 106, 7, 103, 188, 44, 228, 13, 120, 202, 126, 151, 12, 100, 3, 225, 183, 231, 203, 60, 55, 88, 66, 4, 80, 215, 0, 96, 89, 68, 113, 97, 87, 138, 180, 3, 163, 101, 120, 116, 160, 192, 161, 81, 159, 203, 69, 33, 230, 40, 58, 27, 52, 251, 215, 69, 248, 198, 74, 183, 238, 165, 175, 141, 248, 60, 114, 178, 192, 165, 188, 44, 9, 100, 22, 128, 192, 127, 238, 73, 209, 18, 81, 252, 109, 52, 224, 222, 247, 179, 179, 46, 206, 93, 102, 142, 119, 193, 76, 216, 96, 247, 13, 46, 223, 189, 201, 101, 207, 74, 143, 148, 99, 183, 159, 250, 184, 72, 207, 96, 169, 46, 136, 16, 192, 183, 91, 61, 94, 233, 140, 241, 81, 198, 176, 229, 173, 204, 226, 198, 175, 102, 5, 194, 243, 157, 113, 246, 221, 236, 225, 42, 232, 29, 9, 184, 255, 104, 174, 62, 0, 165, 192, 239, 78, 163, 129, 174, 195, 57, 14, 143, 5, 255, 115, 114, 197, 29, 197, 200, 221, 41, 82, 14, 188, 63, 30, 240, 245, 190, 220, 162, 145, 208, 0, 141, 174, 66, 1, 37, 129, 195, 163, 254, 34, 40, 1, 191, 70, 25, 250, 50, 75, 197, 156, 149, 15, 132, 27, 254, 62, 205, 229, 178, 176, 163, 201, 161, 103, 115, 172, 182, 14, 196, 181, 53, 114, 38, 107, 64, 22, 194, 92, 147, 80, 200, 67, 105, 50, 247, 165, 171, 156, 104, 141, 105, 70, 186, 211, 200, 131, 105, 214, 46, 82, 53, 69, 3, 119, 244, 217, 240, 63, 177, 214, 35, 233, 170, 250, 66, 164, 20, 11, 221, 52, 240, 171, 77, 49, 114, 6, 198, 74, 18, 158, 106, 5, 239, 110, 79, 208, 236, 41, 254, 93, 16, 206, 102, 204, 162, 30, 14, 78, 27, 158, 60, 93, 68, 1, 7, 191, 150, 176, 73, 60, 31, 64, 182, 178, 185, 49, 169, 103, 80, 132, 235, 166, 164, 38, 238, 64, 66, 67, 104, 94, 224, 229, 206, 56, 111, 93, 182, 116, 61, 246, 81, 177, 118, 166, 107, 248, 253, 121, 43, 92, 119, 52, 106, 86, 39, 245, 66, 0, 147, 101, 9, 105, 188, 171, 165, 186, 198, 127, 179, 57, 202, 233, 233, 106, 216, 9, 79, 113, 169, 96, 216, 119, 179, 135, 47, 112, 240, 114, 185, 110, 169, 77, 149, 132, 95, 159, 181, 32, 182, 54, 58, 139, 83, 112, 231, 7, 121, 0, 126, 210, 17, 129, 96, 150, 134, 213, 9, 205, 84, 185, 42, 29, 121, 103, 91, 130, 15, 38, 45, 228, 105, 95, 40, 207, 97, 173, 209, 83, 124, 179, 213, 227, 153, 13, 81, 16, 91, 205, 247, 174, 116, 113, 42, 118, 31, 89, 227, 86, 37, 109, 8, 224, 189, 97, 159, 178, 64, 71, 82, 207, 166, 129, 192, 75, 231, 203, 180, 68, 170, 235, 252, 95, 57, 195, 150, 138, 218, 156, 43, 8, 70, 102, 43, 98, 96, 103, 146, 63, 119, 198, 120, 115, 216, 210, 243, 179, 245, 81, 222, 248, 106, 156, 141, 73, 77, 201, 192, 109, 141, 14, 86, 171, 231, 39, 161, 99, 209, 158, 43, 152, 48, 156, 237, 41, 205, 123, 163, 1, 174, 99, 55, 38, 3, 225, 209, 142, 40, 7, 78, 23, 217, 182, 220, 2, 120, 247, 202, 172, 59, 27, 155, 28, 90, 163, 138, 76, 32, 28, 159, 12, 192, 23, 30, 110, 181, 148, 238, 63, 85, 64, 128, 166, 121, 149, 160, 23, 118, 96, 21, 122, 255, 226, 150, 40, 103, 178, 134, 132, 182, 123, 167, 50, 134, 95, 222, 18, 229, 108, 198, 112, 99, 212, 238, 29, 155, 156, 5, 240, 253, 53, 54, 84, 127, 25, 246, 9, 4, 214, 175, 112, 104, 139, 107, 46, 20, 132, 129, 41, 179, 196, 60, 96, 108, 228, 155, 61, 107, 60, 237, 41, 140, 82, 100, 138, 66, 186, 146, 151, 67, 89, 195, 119, 142, 231, 65, 36, 212, 251, 209, 188, 132, 212, 116, 85, 18, 236, 233, 143, 139, 0, 252, 174, 34, 62, 71, 39, 131, 80, 107, 138, 82, 11, 128, 182, 213, 176, 33, 169, 33, 128, 159, 174, 143, 176, 231, 104, 30, 20, 172, 170, 120, 187, 111, 181, 199, 171, 151, 124, 80, 48, 94, 17, 204, 111, 173, 246, 160, 44, 188, 182, 45, 73, 103, 131, 189, 110, 120, 218, 240, 192, 74, 151, 29, 77, 22, 80, 207, 80, 137, 6, 79, 227, 42, 136, 42, 112, 230, 244, 153, 16, 128, 18, 155, 193, 0, 127, 237, 74, 48, 81, 18, 50, 190, 128, 8, 55, 198, 236, 207, 186, 251, 243, 161, 10, 205, 112, 255, 189, 85, 46, 178, 103, 25, 61, 67, 37, 222, 24, 177, 168, 142, 237, 74, 209, 28, 213, 76, 248, 66, 206, 192, 67, 95, 242, 56, 240, 229, 8, 253, 21, 26, 126, 176, 54, 178, 112, 34, 18, 5, 63, 255, 180, 196, 211, 237, 17, 20, 240, 236, 39, 37, 11, 79, 89, 158, 247, 159, 242, 57, 50, 211, 164, 20, 60, 126, 178, 64, 68, 131, 163, 96, 239, 201, 2, 34, 112, 100, 220, 231, 135, 107, 35, 188, 114, 209, 103, 119, 179, 67, 163, 171, 24, 200, 24, 122, 134, 138, 124, 158, 23, 86, 197, 53, 23, 239, 74, 242, 112, 171, 199, 243, 131, 69, 112, 212, 188, 137, 40, 0, 121, 48, 109, 109, 244, 102, 174, 105, 8, 92, 151, 208, 244, 109, 79, 112, 177, 32, 220, 182, 76, 115, 123, 95, 142, 254, 137, 32, 188, 127, 172, 59, 133, 163, 160, 225, 245, 105, 112, 213, 188, 42, 112, 224, 197, 138, 108, 158, 216, 153, 248, 226, 61, 88, 224, 79, 91, 227, 180, 189, 157, 97, 115, 74, 115, 104, 44, 160, 127, 78, 153, 162, 160, 28, 64, 84, 171, 218, 101, 184, 247, 159, 5, 174, 248, 176, 37, 165, 121, 118, 83, 244, 11, 5, 161, 179, 209, 225, 76, 222, 240, 194, 230, 24, 142, 134, 61, 253, 121, 112, 170, 69, 172, 33, 162, 24, 47, 75, 157, 177, 92, 65, 87, 95, 22, 128, 31, 183, 69, 56, 176, 33, 90, 37, 205, 245, 214, 241, 241, 128, 67, 35, 1, 39, 38, 13, 94, 239, 52, 147, 229, 234, 255, 221, 211, 234, 17, 85, 208, 119, 37, 176, 237, 116, 177, 169, 120, 38, 148, 91, 151, 59, 124, 216, 149, 168, 12, 153, 1, 123, 79, 228, 25, 206, 203, 82, 47, 137, 186, 244, 100, 187, 211, 36, 52, 220, 255, 97, 158, 222, 138, 84, 235, 26, 131, 26, 199, 198, 3, 154, 14, 102, 152, 240, 133, 7, 90, 28, 62, 223, 157, 226, 165, 173, 113, 86, 120, 138, 168, 14, 29, 176, 169, 163, 150, 54, 254, 199, 219, 227, 36, 52, 156, 206, 25, 122, 47, 148, 107, 191, 11, 22, 72, 165, 130, 95, 108, 140, 241, 163, 54, 111, 230, 46, 138, 6, 2, 17, 130, 202, 212, 173, 21, 228, 12, 220, 249, 143, 28, 3, 19, 166, 170, 53, 183, 196, 20, 71, 182, 39, 105, 139, 219, 205, 230, 131, 25, 70, 75, 114, 245, 0, 102, 100, 122, 69, 76, 177, 171, 217, 229, 153, 142, 8, 183, 166, 106, 243, 112, 46, 47, 97, 146, 165, 92, 104, 175, 140, 106, 99, 62, 108, 122, 39, 195, 112, 65, 234, 191, 140, 150, 10, 37, 70, 64, 43, 54, 164, 53, 77, 17, 133, 8, 92, 42, 26, 118, 44, 119, 121, 170, 61, 66, 103, 186, 26, 220, 80, 78, 120, 238, 179, 18, 47, 12, 150, 170, 43, 226, 154, 0, 92, 197, 155, 0, 20, 237, 203, 172, 238, 127, 50, 101, 108, 239, 175, 147, 36, 238, 117, 125, 234, 86, 12, 125, 58, 51, 100, 106, 150, 124, 36, 254, 23, 153, 41, 93, 205, 81, 212, 105, 60, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
$objForm.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
$objForm.KeyPreview = $True
$objForm.TabStop = $false


# Add the listbox containing the Get-CsCallParkOrbit cmdlets ============================================================
$objOrbitsListbox = New-Object System.Windows.Forms.Listbox 
$objOrbitsListbox.Location = New-Object System.Drawing.Size(20,30) 
$objOrbitsListbox.Size = New-Object System.Drawing.Size(150,300) 
$objOrbitsListbox.Sorted = $true
$objOrbitsListbox.tabIndex = 10
$objOrbitsListbox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objOrbitsListbox.TabStop = $false

# Add Lync Orbits ============================================================
Get-CsCallParkOrbit -type GroupPickup | select-object Identity | ForEach-Object {[void] $objOrbitsListbox.Items.Add(([string]$_.Identity))}

$objForm.Controls.Add($objOrbitsListbox) 

# Orbits Click Event ============================================================
$objOrbitsListbox.add_Click(
{
	UpdateSelectedOrbit
})

# Orbits Key Event ============================================================
$objOrbitsListbox.add_KeyUp(
{
	if ($_.KeyCode -eq "Up" -or $_.KeyCode -eq "Down") 
	{	
		UpdateSelectedOrbit
	}
})

$objOrbitsLabel = New-Object System.Windows.Forms.Label
$objOrbitsLabel.Location = New-Object System.Drawing.Size(20,15) 
$objOrbitsLabel.Size = New-Object System.Drawing.Size(150,15) 
$objOrbitsLabel.Text = "Call Pickup Orbits"
$objOrbitsLabel.TabStop = $False
$objOrbitsLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($objOrbitsLabel)

<#
[System.Drawing.Pen] $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Black, 1)
[System.Drawing.Graphics] $formGraphics = $objForm.CreateGraphics()
$formGraphics.Drawline($pen, 190,30,190,500)
$pen.Dispose()
$formGraphics.Dispose()
#>

# Add a groupbox ============================================================
$GroupsBox = New-Object System.Windows.Forms.Groupbox
$GroupsBox.Location = New-Object System.Drawing.Size(182,18) 
$GroupsBox.Size = New-Object System.Drawing.Size(1,470) 
$GroupsBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$GroupsBox.TabStop = $False
$GroupsBox.BackColor = [System.Drawing.Color]::DarkGray
$GroupsBox.ForeColor = [System.Drawing.Color]::DarkGray
$objForm.Controls.Add($GroupsBox)

#CHANGED THE WAY GROUPS WORK IN 2.0. THE LISTBOX HAS BEEN REPLACED WITH A MORE FEATURE RICH LIST VIEW.
<#
# Add the listbox containing the Call Pickup Groups ============================================================
$objGroupsListbox = New-Object System.Windows.Forms.Listbox 
$objGroupsListbox.Location = New-Object System.Drawing.Size(195,30) 
$objGroupsListbox.Size = New-Object System.Drawing.Size(150,300) 
$objGroupsListbox.Sorted = $true
$objGroupsListbox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objGroupsListbox.TabStop = $False

# Add Call Pickup Groups Listbox ============================================================
$objForm.Controls.Add($objGroupsListbox) 

# Groups Click Event ============================================================
$objGroupsListbox.add_Click(
{
	UpdateUsersList
	$itemsInUsersList = $objUsersListbox.Items.Count
	if($itemsInUsersList -gt 0)
	{$objUsersListbox.SelectedIndex = 0}
})
# Groups Key Event ============================================================
$objGroupsListbox.add_KeyUp(
{
	if ($_.KeyCode -eq "Up" -or $_.KeyCode -eq "Down") 
	{	
		UpdateUsersList
		$itemsInUsersList = $objUsersListbox.Items.Count
		if($itemsInUsersList -gt 0)
		{$objUsersListbox.SelectedIndex = 0}
	}
})
#>

$objGroupsListbox = New-Object windows.forms.ListView
$objGroupsListbox.View = [System.Windows.Forms.View]"Details"
$objGroupsListbox.Size = New-Object System.Drawing.Size(150,420)
$objGroupsListbox.Location = New-Object System.Drawing.Size(195,30)
$objGroupsListbox.FullRowSelect = $true
$objGroupsListbox.GridLines = $true
$objGroupsListbox.HideSelection = $false
$objGroupsListbox.MultiSelect = $false
$objGroupsListbox.HeaderStyle =  [System.Windows.Forms.ColumnHeaderStyle]::None
$objGroupsListbox.Sorting = [System.Windows.Forms.SortOrder]"Ascending"
$objGroupsListbox.tabIndex = 9
$objGroupsListbox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
[void]$objGroupsListbox.Columns.Add("Groups", 145)
$objForm.Controls.Add($objGroupsListbox)

$objGroupsListbox.add_MouseUp(
{
	UpdateUsersList
	$itemsInUsersList = $objUsersListbox.Items.Count
	if($itemsInUsersList -gt 0)
	{
		$objUsersListbox.SelectedIndex = 0
	}
})

# Groups Key Event ============================================================
$objGroupsListbox.add_KeyUp(
{
	if ($_.KeyCode -eq "Up" -or $_.KeyCode -eq "Down") 
	{	
		UpdateUsersList
		$itemsInUsersList = $objUsersListbox.Items.Count
		if($itemsInUsersList -gt 0)
		{
			$objUsersListbox.SelectedIndex = 0
		}
	}
})

$objGroupsLabel = New-Object System.Windows.Forms.Label
$objGroupsLabel.Location = New-Object System.Drawing.Size(195,15) 
$objGroupsLabel.Size = New-Object System.Drawing.Size(150,15) 
$objGroupsLabel.Text = "Call Pickup Groups"
$objGroupsLabel.TabStop = $false
$objGroupsLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($objGroupsLabel)


# Add the listbox containing the Users in Group ============================================================

$objUsersListbox = New-Object System.Windows.Forms.Listbox 
$objUsersListbox.Location = New-Object System.Drawing.Size(360,30) 
$objUsersListbox.Size = New-Object System.Drawing.Size(230,140) 
$objUsersListbox.SelectionMode = "MultiExtended"
#$objUsersListbox.HorizontalScrollbar = $true
$objUsersListbox.Sorted = $true
$objUsersListbox.tabIndex = 11
$objUsersListbox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Left
$objUsersListbox.TabStop = $false

# Add Users Listbox ============================================================
$objForm.Controls.Add($objUsersListbox) 

<#
$objUsersListbox.add_MouseUp(
{
	#DO NOTHING
})

# User Key Event ============================================================
$objUsersListbox.add_KeyUp(
{
	if ($_.KeyCode -eq "Up" -or $_.KeyCode -eq "Down") 
	{	
		#DO NOTHING
	}
})
#>

$objUsersLabel = New-Object System.Windows.Forms.Label
$objUsersLabel.Location = New-Object System.Drawing.Size(360,15) 
$objUsersLabel.Size = New-Object System.Drawing.Size(150,15) 
$objUsersLabel.Text = "Users in Group"
$objUsersLabel.TabStop = $false
$objUsersLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($objUsersLabel)


# Add the listbox containing the Lync Users ============================================================
$objLyncUsersListbox = New-Object System.Windows.Forms.Listbox 
$objLyncUsersListbox.Location = New-Object System.Drawing.Size(360,215) 
$objLyncUsersListbox.Size = New-Object System.Drawing.Size(230,240) 
$objLyncUsersListbox.SelectionMode = "MultiExtended"
#$objLyncUsersListbox.HorizontalScrollbar = $true
$objLyncUsersListbox.Sorted = $True
$objLyncUsersListbox.tabIndex = 12
$objLyncUsersListbox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objLyncUsersListbox.TabStop = $false

$objForm.Controls.Add($objLyncUsersListbox) 


$objLyncUsersLabel = New-Object System.Windows.Forms.Label
$objLyncUsersLabel.Location = New-Object System.Drawing.Size(360,200) 
$objLyncUsersLabel.Size = New-Object System.Drawing.Size(160,15) 
$objLyncUsersLabel.Text = "Enterprise Voice Users:"
$objLyncUsersLabel.TabStop = $false
$objLyncUsersLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($objLyncUsersLabel)


# Filter button ============================================================
$FilterButton = New-Object System.Windows.Forms.Button
$FilterButton.Location = New-Object System.Drawing.Size(525,457)
$FilterButton.Size = New-Object System.Drawing.Size(60,20)
$FilterButton.Text = "Filter"
$FilterButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$FilterButton.Add_Click({FilterLyncUsersList})
$objForm.Controls.Add($FilterButton)

#Group Number Text box ============================================================
$FilterTextBox = new-object System.Windows.Forms.textbox
$FilterTextBox.location = new-object system.drawing.size(360,457)
$FilterTextBox.size= new-object system.drawing.size(160,15)
$FilterTextBox.text = ""
$FilterTextBox.TabIndex = 8
$FilterTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$FilterTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		FilterLyncUsersList
	}
})
$objform.controls.add($FilterTextBox)


# Find User button ============================================================
$FindUserButton = New-Object System.Windows.Forms.Button
$FindUserButton.Location = New-Object System.Drawing.Size(400,482)
$FindUserButton.Size = New-Object System.Drawing.Size(120,20)
$FindUserButton.Text = "Find Selected User"
$FindUserButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$FindUserButton.Add_Click(
{
	FindUserGroup -User $objLyncUsersListbox.SelectedItems[0]
})
$objForm.Controls.Add($FindUserButton)

[byte[]]$PNGImage = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 12, 0, 0, 0, 12, 8, 6, 0, 0, 0, 86, 117, 92, 231, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 10, 77, 105, 67, 67, 80, 80, 104, 111, 116, 111, 115, 104, 111, 112, 32, 73, 67, 67, 32, 112, 114, 111, 102, 105, 108, 101, 0, 0, 120, 218, 157, 83, 119, 88, 147, 247, 22, 62, 223, 247, 101, 15, 86, 66, 216, 240, 177, 151, 108, 129, 0, 34, 35, 172, 8, 200, 16, 89, 162, 16, 146, 0, 97, 132, 16, 18, 64, 197, 133, 136, 10, 86, 20, 21, 17, 156, 72, 85, 196, 130, 213, 10, 72, 157, 136, 226, 160, 40, 184, 103, 65, 138, 136, 90, 139, 85, 92, 56, 238, 31, 220, 167, 181, 125, 122, 239, 237, 237, 251, 215, 251, 188, 231, 156, 231, 252, 206, 121, 207, 15, 128, 17, 18, 38, 145, 230, 162, 106, 0, 57, 82, 133, 60, 58, 216, 31, 143, 79, 72, 196, 201, 189, 128, 2, 21, 72, 224, 4, 32, 16, 230, 203, 194, 103, 5, 197, 0, 0, 240, 3, 121, 120, 126, 116, 176, 63, 252, 1, 175, 111, 0, 2, 0, 112, 213, 46, 36, 18, 199, 225, 255, 131, 186, 80, 38, 87, 0, 32, 145, 0, 224, 34, 18, 231, 11, 1, 144, 82, 0, 200, 46, 84, 200, 20, 0, 200, 24, 0, 176, 83, 179, 100, 10, 0, 148, 0, 0, 108, 121, 124, 66, 34, 0, 170, 13, 0, 236, 244, 73, 62, 5, 0, 216, 169, 147, 220, 23, 0, 216, 162, 28, 169, 8, 0, 141, 1, 0, 153, 40, 71, 36, 2, 64, 187, 0, 96, 85, 129, 82, 44, 2, 192, 194, 0, 160, 172, 64, 34, 46, 4, 192, 174, 1, 128, 89, 182, 50, 71, 2, 128, 189, 5, 0, 118, 142, 88, 144, 15, 64, 96, 0, 128, 153, 66, 44, 204, 0, 32, 56, 2, 0, 67, 30, 19, 205, 3, 32, 76, 3, 160, 48, 210, 191, 224, 169, 95, 112, 133, 184, 72, 1, 0, 192, 203, 149, 205, 151, 75, 210, 51, 20, 184, 149, 208, 26, 119, 242, 240, 224, 226, 33, 226, 194, 108, 177, 66, 97, 23, 41, 16, 102, 9, 228, 34, 156, 151, 155, 35, 19, 72, 231, 3, 76, 206, 12, 0, 0, 26, 249, 209, 193, 254, 56, 63, 144, 231, 230, 228, 225, 230, 102, 231, 108, 239, 244, 197, 162, 254, 107, 240, 111, 34, 62, 33, 241, 223, 254, 188, 140, 2, 4, 0, 16, 78, 207, 239, 218, 95, 229, 229, 214, 3, 112, 199, 1, 176, 117, 191, 107, 169, 91, 0, 218, 86, 0, 104, 223, 249, 93, 51, 219, 9, 160, 90, 10, 208, 122, 249, 139, 121, 56, 252, 64, 30, 158, 161, 80, 200, 60, 29, 28, 10, 11, 11, 237, 37, 98, 161, 189, 48, 227, 139, 62, 255, 51, 225, 111, 224, 139, 126, 246, 252, 64, 30, 254, 219, 122, 240, 0, 113, 154, 64, 153, 173, 192, 163, 131, 253, 113, 97, 110, 118, 174, 82, 142, 231, 203, 4, 66, 49, 110, 247, 231, 35, 254, 199, 133, 127, 253, 142, 41, 209, 226, 52, 177, 92, 44, 21, 138, 241, 88, 137, 184, 80, 34, 77, 199, 121, 185, 82, 145, 68, 33, 201, 149, 226, 18, 233, 127, 50, 241, 31, 150, 253, 9, 147, 119, 13, 0, 172, 134, 79, 192, 78, 182, 7, 181, 203, 108, 192, 126, 238, 1, 2, 139, 14, 88, 210, 118, 0, 64, 126, 243, 45, 140, 26, 11, 145, 0, 16, 103, 52, 50, 121, 247, 0, 0, 147, 191, 249, 143, 64, 43, 1, 0, 205, 151, 164, 227, 0, 0, 188, 232, 24, 92, 168, 148, 23, 76, 198, 8, 0, 0, 68, 160, 129, 42, 176, 65, 7, 12, 193, 20, 172, 192, 14, 156, 193, 29, 188, 192, 23, 2, 97, 6, 68, 64, 12, 36, 192, 60, 16, 66, 6, 228, 128, 28, 10, 161, 24, 150, 65, 25, 84, 192, 58, 216, 4, 181, 176, 3, 26, 160, 17, 154, 225, 16, 180, 193, 49, 56, 13, 231, 224, 18, 92, 129, 235, 112, 23, 6, 96, 24, 158, 194, 24, 188, 134, 9, 4, 65, 200, 8, 19, 97, 33, 58, 136, 17, 98, 142, 216, 34, 206, 8, 23, 153, 142, 4, 34, 97, 72, 52, 146, 128, 164, 32, 233, 136, 20, 81, 34, 197, 200, 114, 164, 2, 169, 66, 106, 145, 93, 72, 35, 242, 45, 114, 20, 57, 141, 92, 64, 250, 144, 219, 200, 32, 50, 138, 252, 138, 188, 71, 49, 148, 129, 178, 81, 3, 212, 2, 117, 64, 185, 168, 31, 26, 138, 198, 160, 115, 209, 116, 52, 15, 93, 128, 150, 162, 107, 209, 26, 180, 30, 61, 128, 182, 162, 167, 209, 75, 232, 117, 116, 0, 125, 138, 142, 99, 128, 209, 49, 14, 102, 140, 217, 97, 92, 140, 135, 69, 96, 137, 88, 26, 38, 199, 22, 99, 229, 88, 53, 86, 143, 53, 99, 29, 88, 55, 118, 21, 27, 192, 158, 97, 239, 8, 36, 2, 139, 128, 19, 236, 8, 94, 132, 16, 194, 108, 130, 144, 144, 71, 88, 76, 88, 67, 168, 37, 236, 35, 180, 18, 186, 8, 87, 9, 131, 132, 49, 194, 39, 34, 147, 168, 79, 180, 37, 122, 18, 249, 196, 120, 98, 58, 177, 144, 88, 70, 172, 38, 238, 33, 30, 33, 158, 37, 94, 39, 14, 19, 95, 147, 72, 36, 14, 201, 146, 228, 78, 10, 33, 37, 144, 50, 73, 11, 73, 107, 72, 219, 72, 45, 164, 83, 164, 62, 210, 16, 105, 156, 76, 38, 235, 144, 109, 201, 222, 228, 8, 178, 128, 172, 32, 151, 145, 183, 144, 15, 144, 79, 146, 251, 201, 195, 228, 183, 20, 58, 197, 136, 226, 76, 9, 162, 36, 82, 164, 148, 18, 74, 53, 101, 63, 229, 4, 165, 159, 50, 66, 153, 160, 170, 81, 205, 169, 158, 212, 8, 170, 136, 58, 159, 90, 73, 109, 160, 118, 80, 47, 83, 135, 169, 19, 52, 117, 154, 37, 205, 155, 22, 67, 203, 164, 45, 163, 213, 208, 154, 105, 103, 105, 247, 104, 47, 233, 116, 186, 9, 221, 131, 30, 69, 151, 208, 151, 210, 107, 232, 7, 233, 231, 233, 131, 244, 119, 12, 13, 134, 13, 131, 199, 72, 98, 40, 25, 107, 25, 123, 25, 167, 24, 183, 25, 47, 153, 76, 166, 5, 211, 151, 153, 200, 84, 48, 215, 50, 27, 153, 103, 152, 15, 152, 111, 85, 88, 42, 246, 42, 124, 21, 145, 202, 18, 149, 58, 149, 86, 149, 126, 149, 231, 170, 84, 85, 115, 85, 63, 213, 121, 170, 11, 84, 171, 85, 15, 171, 94, 86, 125, 166, 70, 85, 179, 80, 227, 169, 9, 212, 22, 171, 213, 169, 29, 85, 187, 169, 54, 174, 206, 82, 119, 82, 143, 80, 207, 81, 95, 163, 190, 95, 253, 130, 250, 99, 13, 178, 134, 133, 70, 160, 134, 72, 163, 84, 99, 183, 198, 25, 141, 33, 22, 198, 50, 101, 241, 88, 66, 214, 114, 86, 3, 235, 44, 107, 152, 77, 98, 91, 178, 249, 236, 76, 118, 5, 251, 27, 118, 47, 123, 76, 83, 67, 115, 170, 102, 172, 102, 145, 102, 157, 230, 113, 205, 1, 14, 198, 177, 224, 240, 57, 217, 156, 74, 206, 33, 206, 13, 206, 123, 45, 3, 45, 63, 45, 177, 214, 106, 173, 102, 173, 126, 173, 55, 218, 122, 218, 190, 218, 98, 237, 114, 237, 22, 237, 235, 218, 239, 117, 112, 157, 64, 157, 44, 157, 245, 58, 109, 58, 247, 117, 9, 186, 54, 186, 81, 186, 133, 186, 219, 117, 207, 234, 62, 211, 99, 235, 121, 233, 9, 245, 202, 245, 14, 233, 221, 209, 71, 245, 109, 244, 163, 245, 23, 234, 239, 214, 239, 209, 31, 55, 48, 52, 8, 54, 144, 25, 108, 49, 56, 99, 240, 204, 144, 99, 232, 107, 152, 105, 184, 209, 240, 132, 225, 168, 17, 203, 104, 186, 145, 196, 104, 163, 209, 73, 163, 39, 184, 38, 238, 135, 103, 227, 53, 120, 23, 62, 102, 172, 111, 28, 98, 172, 52, 222, 101, 220, 107, 60, 97, 98, 105, 50, 219, 164, 196, 164, 197, 228, 190, 41, 205, 148, 107, 154, 102, 186, 209, 180, 211, 116, 204, 204, 200, 44, 220, 172, 216, 172, 201, 236, 142, 57, 213, 156, 107, 158, 97, 190, 217, 188, 219, 252, 141, 133, 165, 69, 156, 197, 74, 139, 54, 139, 199, 150, 218, 150, 124, 203, 5, 150, 77, 150, 247, 172, 152, 86, 62, 86, 121, 86, 245, 86, 215, 172, 73, 214, 92, 235, 44, 235, 109, 214, 87, 108, 80, 27, 87, 155, 12, 155, 58, 155, 203, 182, 168, 173, 155, 173, 196, 118, 155, 109, 223, 20, 226, 20, 143, 41, 210, 41, 245, 83, 110, 218, 49, 236, 252, 236, 10, 236, 154, 236, 6, 237, 57, 246, 97, 246, 37, 246, 109, 246, 207, 29, 204, 28, 18, 29, 214, 59, 116, 59, 124, 114, 116, 117, 204, 118, 108, 112, 188, 235, 164, 225, 52, 195, 169, 196, 169, 195, 233, 87, 103, 27, 103, 161, 115, 157, 243, 53, 23, 166, 75, 144, 203, 18, 151, 118, 151, 23, 83, 109, 167, 138, 167, 110, 159, 122, 203, 149, 229, 26, 238, 186, 210, 181, 211, 245, 163, 155, 187, 155, 220, 173, 217, 109, 212, 221, 204, 61, 197, 125, 171, 251, 77, 46, 155, 27, 201, 93, 195, 61, 239, 65, 244, 240, 247, 88, 226, 113, 204, 227, 157, 167, 155, 167, 194, 243, 144, 231, 47, 94, 118, 94, 89, 94, 251, 189, 30, 79, 179, 156, 38, 158, 214, 48, 109, 200, 219, 196, 91, 224, 189, 203, 123, 96, 58, 62, 61, 101, 250, 206, 233, 3, 62, 198, 62, 2, 159, 122, 159, 135, 190, 166, 190, 34, 223, 61, 190, 35, 126, 214, 126, 153, 126, 7, 252, 158, 251, 59, 250, 203, 253, 143, 248, 191, 225, 121, 242, 22, 241, 78, 5, 96, 1, 193, 1, 229, 1, 189, 129, 26, 129, 179, 3, 107, 3, 31, 4, 153, 4, 165, 7, 53, 5, 141, 5, 187, 6, 47, 12, 62, 21, 66, 12, 9, 13, 89, 31, 114, 147, 111, 192, 23, 242, 27, 249, 99, 51, 220, 103, 44, 154, 209, 21, 202, 8, 157, 21, 90, 27, 250, 48, 204, 38, 76, 30, 214, 17, 142, 134, 207, 8, 223, 16, 126, 111, 166, 249, 76, 233, 204, 182, 8, 136, 224, 71, 108, 136, 184, 31, 105, 25, 153, 23, 249, 125, 20, 41, 42, 50, 170, 46, 234, 81, 180, 83, 116, 113, 116, 247, 44, 214, 172, 228, 89, 251, 103, 189, 142, 241, 143, 169, 140, 185, 59, 219, 106, 182, 114, 118, 103, 172, 106, 108, 82, 108, 99, 236, 155, 184, 128, 184, 170, 184, 129, 120, 135, 248, 69, 241, 151, 18, 116, 19, 36, 9, 237, 137, 228, 196, 216, 196, 61, 137, 227, 115, 2, 231, 108, 154, 51, 156, 228, 154, 84, 150, 116, 99, 174, 229, 220, 162, 185, 23, 230, 233, 206, 203, 158, 119, 60, 89, 53, 89, 144, 124, 56, 133, 152, 18, 151, 178, 63, 229, 131, 32, 66, 80, 47, 24, 79, 229, 167, 110, 77, 29, 19, 242, 132, 155, 133, 79, 69, 190, 162, 141, 162, 81, 177, 183, 184, 74, 60, 146, 230, 157, 86, 149, 246, 56, 221, 59, 125, 67, 250, 104, 134, 79, 70, 117, 198, 51, 9, 79, 82, 43, 121, 145, 25, 146, 185, 35, 243, 77, 86, 68, 214, 222, 172, 207, 217, 113, 217, 45, 57, 148, 156, 148, 156, 163, 82, 13, 105, 150, 180, 43, 215, 48, 183, 40, 183, 79, 102, 43, 43, 147, 13, 228, 121, 230, 109, 202, 27, 147, 135, 202, 247, 228, 35, 249, 115, 243, 219, 21, 108, 133, 76, 209, 163, 180, 82, 174, 80, 14, 22, 76, 47, 168, 43, 120, 91, 24, 91, 120, 184, 72, 189, 72, 90, 212, 51, 223, 102, 254, 234, 249, 35, 11, 130, 22, 124, 189, 144, 176, 80, 184, 176, 179, 216, 184, 120, 89, 241, 224, 34, 191, 69, 187, 22, 35, 139, 83, 23, 119, 46, 49, 93, 82, 186, 100, 120, 105, 240, 210, 125, 203, 104, 203, 178, 150, 253, 80, 226, 88, 82, 85, 242, 106, 121, 220, 242, 142, 82, 131, 210, 165, 165, 67, 43, 130, 87, 52, 149, 169, 148, 201, 203, 110, 174, 244, 90, 185, 99, 21, 97, 149, 100, 85, 239, 106, 151, 213, 91, 86, 127, 42, 23, 149, 95, 172, 112, 172, 168, 174, 248, 176, 70, 184, 230, 226, 87, 78, 95, 213, 124, 245, 121, 109, 218, 218, 222, 74, 183, 202, 237, 235, 72, 235, 164, 235, 110, 172, 247, 89, 191, 175, 74, 189, 106, 65, 213, 208, 134, 240, 13, 173, 27, 241, 141, 229, 27, 95, 109, 74, 222, 116, 161, 122, 106, 245, 142, 205, 180, 205, 202, 205, 3, 53, 97, 53, 237, 91, 204, 182, 172, 219, 242, 161, 54, 163, 246, 122, 157, 127, 93, 203, 86, 253, 173, 171, 183, 190, 217, 38, 218, 214, 191, 221, 119, 123, 243, 14, 131, 29, 21, 59, 222, 239, 148, 236, 188, 181, 43, 120, 87, 107, 189, 69, 125, 245, 110, 210, 238, 130, 221, 143, 26, 98, 27, 186, 191, 230, 126, 221, 184, 71, 119, 79, 197, 158, 143, 123, 165, 123, 7, 246, 69, 239, 235, 106, 116, 111, 108, 220, 175, 191, 191, 178, 9, 109, 82, 54, 141, 30, 72, 58, 112, 229, 155, 128, 111, 218, 155, 237, 154, 119, 181, 112, 90, 42, 14, 194, 65, 229, 193, 39, 223, 166, 124, 123, 227, 80, 232, 161, 206, 195, 220, 195, 205, 223, 153, 127, 183, 245, 8, 235, 72, 121, 43, 210, 58, 191, 117, 172, 45, 163, 109, 160, 61, 161, 189, 239, 232, 140, 163, 157, 29, 94, 29, 71, 190, 183, 255, 126, 239, 49, 227, 99, 117, 199, 53, 143, 87, 158, 160, 157, 40, 61, 241, 249, 228, 130, 147, 227, 167, 100, 167, 158, 157, 78, 63, 61, 212, 153, 220, 121, 247, 76, 252, 153, 107, 93, 81, 93, 189, 103, 67, 207, 158, 63, 23, 116, 238, 76, 183, 95, 247, 201, 243, 222, 231, 143, 93, 240, 188, 112, 244, 34, 247, 98, 219, 37, 183, 75, 173, 61, 174, 61, 71, 126, 112, 253, 225, 72, 175, 91, 111, 235, 101, 247, 203, 237, 87, 60, 174, 116, 244, 77, 235, 59, 209, 239, 211, 127, 250, 106, 192, 213, 115, 215, 248, 215, 46, 93, 159, 121, 189, 239, 198, 236, 27, 183, 110, 38, 221, 28, 184, 37, 186, 245, 248, 118, 246, 237, 23, 119, 10, 238, 76, 220, 93, 122, 143, 120, 175, 252, 190, 218, 253, 234, 7, 250, 15, 234, 127, 180, 254, 177, 101, 192, 109, 224, 248, 96, 192, 96, 207, 195, 89, 15, 239, 14, 9, 135, 158, 254, 148, 255, 211, 135, 225, 210, 71, 204, 71, 213, 35, 70, 35, 141, 143, 157, 31, 31, 27, 13, 26, 189, 242, 100, 206, 147, 225, 167, 178, 167, 19, 207, 202, 126, 86, 255, 121, 235, 115, 171, 231, 223, 253, 226, 251, 75, 207, 88, 252, 216, 240, 11, 249, 139, 207, 191, 174, 121, 169, 243, 114, 239, 171, 169, 175, 58, 199, 35, 199, 31, 188, 206, 121, 61, 241, 166, 252, 173, 206, 219, 125, 239, 184, 239, 186, 223, 199, 189, 31, 153, 40, 252, 64, 254, 80, 243, 209, 250, 99, 199, 167, 208, 79, 247, 62, 231, 124, 254, 252, 47, 247, 132, 243, 251, 37, 210, 159, 51, 0, 0, 0, 4, 103, 65, 77, 65, 0, 0, 177, 142, 124, 251, 81, 147, 0, 0, 0, 32, 99, 72, 82, 77, 0, 0, 122, 37, 0, 0, 128, 131, 0, 0, 249, 255, 0, 0, 128, 233, 0, 0, 117, 48, 0, 0, 234, 96, 0, 0, 58, 152, 0, 0, 23, 111, 146, 95, 197, 70, 0, 0, 0, 131, 73, 68, 65, 84, 120, 218, 164, 209, 59, 10, 194, 80, 16, 133, 225, 47, 226, 163, 245, 209, 90, 184, 65, 55, 97, 208, 37, 8, 22, 182, 118, 22, 238, 197, 86, 236, 45, 109, 132, 64, 184, 54, 19, 72, 145, 24, 46, 30, 248, 97, 96, 254, 3, 3, 83, 164, 148, 252, 155, 2, 59, 148, 49, 15, 102, 131, 123, 176, 30, 146, 151, 184, 33, 5, 87, 172, 250, 228, 41, 206, 45, 185, 225, 136, 89, 215, 221, 251, 14, 185, 161, 108, 203, 115, 108, 81, 253, 40, 124, 194, 89, 192, 9, 239, 88, 84, 61, 36, 188, 112, 24, 227, 129, 11, 234, 160, 43, 147, 40, 62, 139, 220, 199, 141, 114, 191, 154, 93, 248, 14, 0, 33, 182, 49, 198, 152, 20, 2, 249, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
$PNGStream = New-Object IO.MemoryStream($PNGImage, 0, $PNGImage.Length)
[System.Drawing.Image] $theImage = [System.Drawing.Image]::FromStream($PNGStream)


#Add User button ============================================================
$addUserButton = New-Object System.Windows.Forms.Button
$addUserButton.Location = New-Object System.Drawing.Size(360,170)
$addUserButton.Size = New-Object System.Drawing.Size(110,23)
$addUserButton.Text = "Add User(s)"
$addUserButton.Image = $theImage
$addUserButton.TextImageRelation = [System.Windows.Forms.TextImageRelation]::TextBeforeImage
$addUserButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$addUserButton.Add_Click(
{
	$StatusLabel.Text = ""
	[string] $listSelectedItem = $objGroupsListbox.SelectedItems[0].Text
	if($listSelectedItem -ne "")
	{
	$addUserButton.enabled = $false
	$deleteUserButton.enabled = $false
	$addOrbitButton.enabled = $false
	$deleteOrbitButton.enabled = $false
	$refreshButton.enabled = $false
	$FilterButton.enabled = $false
	$FindUserButton.enabled = $false
	$exportSCVButton.Enabled = $false
	
	$StatusLabel.Text = "Please Wait: Adding User(s)"
	[System.Windows.Forms.Application]::DoEvents()
	Write-Host ""
	Write-Host "Adding user(s) to group..." -foreground "yellow"
	Write-Host ""
	if($Script:SkypeForBusinessAvailable)
	{
		AddUserToGroupSkype
	}
	else
	{
		AddUserToGroup
		Start-Sleep -s 1
	}
	
	#THIS CODE CAN BE USED TO REFRESH CONFIGURATION FROM THE SYSTEM AFTER EVERY SETTING CHANGE. HOWEVER, THIS CAN SLOW THINGS DOWN. SO IT'S NOT DONE ANYMORE.
	<#
	if($Script:SkypeForBusinessAvailable)
	{
		LoadGroupsSkype
	}
	else
	{
		LoadGroups
	}
	#>
	#Update Group List
	UpdateGroupsList
	
	$StatusLabel.Text = ""
	UpdateUsersList
	$itemsInUsersList = $objUsersListbox.Items.Count
	if($itemsInUsersList -gt 0)
	{$objUsersListbox.SelectedIndex = 0}
	
	$addUserButton.enabled = $true
	$deleteUserButton.enabled = $true
	$addOrbitButton.enabled = $true
	$deleteOrbitButton.enabled = $true
	$refreshButton.enabled = $true
	$FilterButton.enabled = $true
	$FindUserButton.enabled = $true
	$exportSCVButton.Enabled = $true
	}
	else
	{
		Write-host "ERROR: You need to first create a Group to add the user to..." -foreground "red"
	}
})
$objForm.Controls.Add($addUserButton)


[byte[]]$PNGImage = @(137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 12, 0, 0, 0, 12, 8, 6, 0, 0, 0, 86, 117, 92, 231, 0, 0, 0, 9, 112, 72, 89, 115, 0, 0, 11, 19, 0, 0, 11, 19, 1, 0, 154, 156, 24, 0, 0, 10, 77, 105, 67, 67, 80, 80, 104, 111, 116, 111, 115, 104, 111, 112, 32, 73, 67, 67, 32, 112, 114, 111, 102, 105, 108, 101, 0, 0, 120, 218, 157, 83, 119, 88, 147, 247, 22, 62, 223, 247, 101, 15, 86, 66, 216, 240, 177, 151, 108, 129, 0, 34, 35, 172, 8, 200, 16, 89, 162, 16, 146, 0, 97, 132, 16, 18, 64, 197, 133, 136, 10, 86, 20, 21, 17, 156, 72, 85, 196, 130, 213, 10, 72, 157, 136, 226, 160, 40, 184, 103, 65, 138, 136, 90, 139, 85, 92, 56, 238, 31, 220, 167, 181, 125, 122, 239, 237, 237, 251, 215, 251, 188, 231, 156, 231, 252, 206, 121, 207, 15, 128, 17, 18, 38, 145, 230, 162, 106, 0, 57, 82, 133, 60, 58, 216, 31, 143, 79, 72, 196, 201, 189, 128, 2, 21, 72, 224, 4, 32, 16, 230, 203, 194, 103, 5, 197, 0, 0, 240, 3, 121, 120, 126, 116, 176, 63, 252, 1, 175, 111, 0, 2, 0, 112, 213, 46, 36, 18, 199, 225, 255, 131, 186, 80, 38, 87, 0, 32, 145, 0, 224, 34, 18, 231, 11, 1, 144, 82, 0, 200, 46, 84, 200, 20, 0, 200, 24, 0, 176, 83, 179, 100, 10, 0, 148, 0, 0, 108, 121, 124, 66, 34, 0, 170, 13, 0, 236, 244, 73, 62, 5, 0, 216, 169, 147, 220, 23, 0, 216, 162, 28, 169, 8, 0, 141, 1, 0, 153, 40, 71, 36, 2, 64, 187, 0, 96, 85, 129, 82, 44, 2, 192, 194, 0, 160, 172, 64, 34, 46, 4, 192, 174, 1, 128, 89, 182, 50, 71, 2, 128, 189, 5, 0, 118, 142, 88, 144, 15, 64, 96, 0, 128, 153, 66, 44, 204, 0, 32, 56, 2, 0, 67, 30, 19, 205, 3, 32, 76, 3, 160, 48, 210, 191, 224, 169, 95, 112, 133, 184, 72, 1, 0, 192, 203, 149, 205, 151, 75, 210, 51, 20, 184, 149, 208, 26, 119, 242, 240, 224, 226, 33, 226, 194, 108, 177, 66, 97, 23, 41, 16, 102, 9, 228, 34, 156, 151, 155, 35, 19, 72, 231, 3, 76, 206, 12, 0, 0, 26, 249, 209, 193, 254, 56, 63, 144, 231, 230, 228, 225, 230, 102, 231, 108, 239, 244, 197, 162, 254, 107, 240, 111, 34, 62, 33, 241, 223, 254, 188, 140, 2, 4, 0, 16, 78, 207, 239, 218, 95, 229, 229, 214, 3, 112, 199, 1, 176, 117, 191, 107, 169, 91, 0, 218, 86, 0, 104, 223, 249, 93, 51, 219, 9, 160, 90, 10, 208, 122, 249, 139, 121, 56, 252, 64, 30, 158, 161, 80, 200, 60, 29, 28, 10, 11, 11, 237, 37, 98, 161, 189, 48, 227, 139, 62, 255, 51, 225, 111, 224, 139, 126, 246, 252, 64, 30, 254, 219, 122, 240, 0, 113, 154, 64, 153, 173, 192, 163, 131, 253, 113, 97, 110, 118, 174, 82, 142, 231, 203, 4, 66, 49, 110, 247, 231, 35, 254, 199, 133, 127, 253, 142, 41, 209, 226, 52, 177, 92, 44, 21, 138, 241, 88, 137, 184, 80, 34, 77, 199, 121, 185, 82, 145, 68, 33, 201, 149, 226, 18, 233, 127, 50, 241, 31, 150, 253, 9, 147, 119, 13, 0, 172, 134, 79, 192, 78, 182, 7, 181, 203, 108, 192, 126, 238, 1, 2, 139, 14, 88, 210, 118, 0, 64, 126, 243, 45, 140, 26, 11, 145, 0, 16, 103, 52, 50, 121, 247, 0, 0, 147, 191, 249, 143, 64, 43, 1, 0, 205, 151, 164, 227, 0, 0, 188, 232, 24, 92, 168, 148, 23, 76, 198, 8, 0, 0, 68, 160, 129, 42, 176, 65, 7, 12, 193, 20, 172, 192, 14, 156, 193, 29, 188, 192, 23, 2, 97, 6, 68, 64, 12, 36, 192, 60, 16, 66, 6, 228, 128, 28, 10, 161, 24, 150, 65, 25, 84, 192, 58, 216, 4, 181, 176, 3, 26, 160, 17, 154, 225, 16, 180, 193, 49, 56, 13, 231, 224, 18, 92, 129, 235, 112, 23, 6, 96, 24, 158, 194, 24, 188, 134, 9, 4, 65, 200, 8, 19, 97, 33, 58, 136, 17, 98, 142, 216, 34, 206, 8, 23, 153, 142, 4, 34, 97, 72, 52, 146, 128, 164, 32, 233, 136, 20, 81, 34, 197, 200, 114, 164, 2, 169, 66, 106, 145, 93, 72, 35, 242, 45, 114, 20, 57, 141, 92, 64, 250, 144, 219, 200, 32, 50, 138, 252, 138, 188, 71, 49, 148, 129, 178, 81, 3, 212, 2, 117, 64, 185, 168, 31, 26, 138, 198, 160, 115, 209, 116, 52, 15, 93, 128, 150, 162, 107, 209, 26, 180, 30, 61, 128, 182, 162, 167, 209, 75, 232, 117, 116, 0, 125, 138, 142, 99, 128, 209, 49, 14, 102, 140, 217, 97, 92, 140, 135, 69, 96, 137, 88, 26, 38, 199, 22, 99, 229, 88, 53, 86, 143, 53, 99, 29, 88, 55, 118, 21, 27, 192, 158, 97, 239, 8, 36, 2, 139, 128, 19, 236, 8, 94, 132, 16, 194, 108, 130, 144, 144, 71, 88, 76, 88, 67, 168, 37, 236, 35, 180, 18, 186, 8, 87, 9, 131, 132, 49, 194, 39, 34, 147, 168, 79, 180, 37, 122, 18, 249, 196, 120, 98, 58, 177, 144, 88, 70, 172, 38, 238, 33, 30, 33, 158, 37, 94, 39, 14, 19, 95, 147, 72, 36, 14, 201, 146, 228, 78, 10, 33, 37, 144, 50, 73, 11, 73, 107, 72, 219, 72, 45, 164, 83, 164, 62, 210, 16, 105, 156, 76, 38, 235, 144, 109, 201, 222, 228, 8, 178, 128, 172, 32, 151, 145, 183, 144, 15, 144, 79, 146, 251, 201, 195, 228, 183, 20, 58, 197, 136, 226, 76, 9, 162, 36, 82, 164, 148, 18, 74, 53, 101, 63, 229, 4, 165, 159, 50, 66, 153, 160, 170, 81, 205, 169, 158, 212, 8, 170, 136, 58, 159, 90, 73, 109, 160, 118, 80, 47, 83, 135, 169, 19, 52, 117, 154, 37, 205, 155, 22, 67, 203, 164, 45, 163, 213, 208, 154, 105, 103, 105, 247, 104, 47, 233, 116, 186, 9, 221, 131, 30, 69, 151, 208, 151, 210, 107, 232, 7, 233, 231, 233, 131, 244, 119, 12, 13, 134, 13, 131, 199, 72, 98, 40, 25, 107, 25, 123, 25, 167, 24, 183, 25, 47, 153, 76, 166, 5, 211, 151, 153, 200, 84, 48, 215, 50, 27, 153, 103, 152, 15, 152, 111, 85, 88, 42, 246, 42, 124, 21, 145, 202, 18, 149, 58, 149, 86, 149, 126, 149, 231, 170, 84, 85, 115, 85, 63, 213, 121, 170, 11, 84, 171, 85, 15, 171, 94, 86, 125, 166, 70, 85, 179, 80, 227, 169, 9, 212, 22, 171, 213, 169, 29, 85, 187, 169, 54, 174, 206, 82, 119, 82, 143, 80, 207, 81, 95, 163, 190, 95, 253, 130, 250, 99, 13, 178, 134, 133, 70, 160, 134, 72, 163, 84, 99, 183, 198, 25, 141, 33, 22, 198, 50, 101, 241, 88, 66, 214, 114, 86, 3, 235, 44, 107, 152, 77, 98, 91, 178, 249, 236, 76, 118, 5, 251, 27, 118, 47, 123, 76, 83, 67, 115, 170, 102, 172, 102, 145, 102, 157, 230, 113, 205, 1, 14, 198, 177, 224, 240, 57, 217, 156, 74, 206, 33, 206, 13, 206, 123, 45, 3, 45, 63, 45, 177, 214, 106, 173, 102, 173, 126, 173, 55, 218, 122, 218, 190, 218, 98, 237, 114, 237, 22, 237, 235, 218, 239, 117, 112, 157, 64, 157, 44, 157, 245, 58, 109, 58, 247, 117, 9, 186, 54, 186, 81, 186, 133, 186, 219, 117, 207, 234, 62, 211, 99, 235, 121, 233, 9, 245, 202, 245, 14, 233, 221, 209, 71, 245, 109, 244, 163, 245, 23, 234, 239, 214, 239, 209, 31, 55, 48, 52, 8, 54, 144, 25, 108, 49, 56, 99, 240, 204, 144, 99, 232, 107, 152, 105, 184, 209, 240, 132, 225, 168, 17, 203, 104, 186, 145, 196, 104, 163, 209, 73, 163, 39, 184, 38, 238, 135, 103, 227, 53, 120, 23, 62, 102, 172, 111, 28, 98, 172, 52, 222, 101, 220, 107, 60, 97, 98, 105, 50, 219, 164, 196, 164, 197, 228, 190, 41, 205, 148, 107, 154, 102, 186, 209, 180, 211, 116, 204, 204, 200, 44, 220, 172, 216, 172, 201, 236, 142, 57, 213, 156, 107, 158, 97, 190, 217, 188, 219, 252, 141, 133, 165, 69, 156, 197, 74, 139, 54, 139, 199, 150, 218, 150, 124, 203, 5, 150, 77, 150, 247, 172, 152, 86, 62, 86, 121, 86, 245, 86, 215, 172, 73, 214, 92, 235, 44, 235, 109, 214, 87, 108, 80, 27, 87, 155, 12, 155, 58, 155, 203, 182, 168, 173, 155, 173, 196, 118, 155, 109, 223, 20, 226, 20, 143, 41, 210, 41, 245, 83, 110, 218, 49, 236, 252, 236, 10, 236, 154, 236, 6, 237, 57, 246, 97, 246, 37, 246, 109, 246, 207, 29, 204, 28, 18, 29, 214, 59, 116, 59, 124, 114, 116, 117, 204, 118, 108, 112, 188, 235, 164, 225, 52, 195, 169, 196, 169, 195, 233, 87, 103, 27, 103, 161, 115, 157, 243, 53, 23, 166, 75, 144, 203, 18, 151, 118, 151, 23, 83, 109, 167, 138, 167, 110, 159, 122, 203, 149, 229, 26, 238, 186, 210, 181, 211, 245, 163, 155, 187, 155, 220, 173, 217, 109, 212, 221, 204, 61, 197, 125, 171, 251, 77, 46, 155, 27, 201, 93, 195, 61, 239, 65, 244, 240, 247, 88, 226, 113, 204, 227, 157, 167, 155, 167, 194, 243, 144, 231, 47, 94, 118, 94, 89, 94, 251, 189, 30, 79, 179, 156, 38, 158, 214, 48, 109, 200, 219, 196, 91, 224, 189, 203, 123, 96, 58, 62, 61, 101, 250, 206, 233, 3, 62, 198, 62, 2, 159, 122, 159, 135, 190, 166, 190, 34, 223, 61, 190, 35, 126, 214, 126, 153, 126, 7, 252, 158, 251, 59, 250, 203, 253, 143, 248, 191, 225, 121, 242, 22, 241, 78, 5, 96, 1, 193, 1, 229, 1, 189, 129, 26, 129, 179, 3, 107, 3, 31, 4, 153, 4, 165, 7, 53, 5, 141, 5, 187, 6, 47, 12, 62, 21, 66, 12, 9, 13, 89, 31, 114, 147, 111, 192, 23, 242, 27, 249, 99, 51, 220, 103, 44, 154, 209, 21, 202, 8, 157, 21, 90, 27, 250, 48, 204, 38, 76, 30, 214, 17, 142, 134, 207, 8, 223, 16, 126, 111, 166, 249, 76, 233, 204, 182, 8, 136, 224, 71, 108, 136, 184, 31, 105, 25, 153, 23, 249, 125, 20, 41, 42, 50, 170, 46, 234, 81, 180, 83, 116, 113, 116, 247, 44, 214, 172, 228, 89, 251, 103, 189, 142, 241, 143, 169, 140, 185, 59, 219, 106, 182, 114, 118, 103, 172, 106, 108, 82, 108, 99, 236, 155, 184, 128, 184, 170, 184, 129, 120, 135, 248, 69, 241, 151, 18, 116, 19, 36, 9, 237, 137, 228, 196, 216, 196, 61, 137, 227, 115, 2, 231, 108, 154, 51, 156, 228, 154, 84, 150, 116, 99, 174, 229, 220, 162, 185, 23, 230, 233, 206, 203, 158, 119, 60, 89, 53, 89, 144, 124, 56, 133, 152, 18, 151, 178, 63, 229, 131, 32, 66, 80, 47, 24, 79, 229, 167, 110, 77, 29, 19, 242, 132, 155, 133, 79, 69, 190, 162, 141, 162, 81, 177, 183, 184, 74, 60, 146, 230, 157, 86, 149, 246, 56, 221, 59, 125, 67, 250, 104, 134, 79, 70, 117, 198, 51, 9, 79, 82, 43, 121, 145, 25, 146, 185, 35, 243, 77, 86, 68, 214, 222, 172, 207, 217, 113, 217, 45, 57, 148, 156, 148, 156, 163, 82, 13, 105, 150, 180, 43, 215, 48, 183, 40, 183, 79, 102, 43, 43, 147, 13, 228, 121, 230, 109, 202, 27, 147, 135, 202, 247, 228, 35, 249, 115, 243, 219, 21, 108, 133, 76, 209, 163, 180, 82, 174, 80, 14, 22, 76, 47, 168, 43, 120, 91, 24, 91, 120, 184, 72, 189, 72, 90, 212, 51, 223, 102, 254, 234, 249, 35, 11, 130, 22, 124, 189, 144, 176, 80, 184, 176, 179, 216, 184, 120, 89, 241, 224, 34, 191, 69, 187, 22, 35, 139, 83, 23, 119, 46, 49, 93, 82, 186, 100, 120, 105, 240, 210, 125, 203, 104, 203, 178, 150, 253, 80, 226, 88, 82, 85, 242, 106, 121, 220, 242, 142, 82, 131, 210, 165, 165, 67, 43, 130, 87, 52, 149, 169, 148, 201, 203, 110, 174, 244, 90, 185, 99, 21, 97, 149, 100, 85, 239, 106, 151, 213, 91, 86, 127, 42, 23, 149, 95, 172, 112, 172, 168, 174, 248, 176, 70, 184, 230, 226, 87, 78, 95, 213, 124, 245, 121, 109, 218, 218, 222, 74, 183, 202, 237, 235, 72, 235, 164, 235, 110, 172, 247, 89, 191, 175, 74, 189, 106, 65, 213, 208, 134, 240, 13, 173, 27, 241, 141, 229, 27, 95, 109, 74, 222, 116, 161, 122, 106, 245, 142, 205, 180, 205, 202, 205, 3, 53, 97, 53, 237, 91, 204, 182, 172, 219, 242, 161, 54, 163, 246, 122, 157, 127, 93, 203, 86, 253, 173, 171, 183, 190, 217, 38, 218, 214, 191, 221, 119, 123, 243, 14, 131, 29, 21, 59, 222, 239, 148, 236, 188, 181, 43, 120, 87, 107, 189, 69, 125, 245, 110, 210, 238, 130, 221, 143, 26, 98, 27, 186, 191, 230, 126, 221, 184, 71, 119, 79, 197, 158, 143, 123, 165, 123, 7, 246, 69, 239, 235, 106, 116, 111, 108, 220, 175, 191, 191, 178, 9, 109, 82, 54, 141, 30, 72, 58, 112, 229, 155, 128, 111, 218, 155, 237, 154, 119, 181, 112, 90, 42, 14, 194, 65, 229, 193, 39, 223, 166, 124, 123, 227, 80, 232, 161, 206, 195, 220, 195, 205, 223, 153, 127, 183, 245, 8, 235, 72, 121, 43, 210, 58, 191, 117, 172, 45, 163, 109, 160, 61, 161, 189, 239, 232, 140, 163, 157, 29, 94, 29, 71, 190, 183, 255, 126, 239, 49, 227, 99, 117, 199, 53, 143, 87, 158, 160, 157, 40, 61, 241, 249, 228, 130, 147, 227, 167, 100, 167, 158, 157, 78, 63, 61, 212, 153, 220, 121, 247, 76, 252, 153, 107, 93, 81, 93, 189, 103, 67, 207, 158, 63, 23, 116, 238, 76, 183, 95, 247, 201, 243, 222, 231, 143, 93, 240, 188, 112, 244, 34, 247, 98, 219, 37, 183, 75, 173, 61, 174, 61, 71, 126, 112, 253, 225, 72, 175, 91, 111, 235, 101, 247, 203, 237, 87, 60, 174, 116, 244, 77, 235, 59, 209, 239, 211, 127, 250, 106, 192, 213, 115, 215, 248, 215, 46, 93, 159, 121, 189, 239, 198, 236, 27, 183, 110, 38, 221, 28, 184, 37, 186, 245, 248, 118, 246, 237, 23, 119, 10, 238, 76, 220, 93, 122, 143, 120, 175, 252, 190, 218, 253, 234, 7, 250, 15, 234, 127, 180, 254, 177, 101, 192, 109, 224, 248, 96, 192, 96, 207, 195, 89, 15, 239, 14, 9, 135, 158, 254, 148, 255, 211, 135, 225, 210, 71, 204, 71, 213, 35, 70, 35, 141, 143, 157, 31, 31, 27, 13, 26, 189, 242, 100, 206, 147, 225, 167, 178, 167, 19, 207, 202, 126, 86, 255, 121, 235, 115, 171, 231, 223, 253, 226, 251, 75, 207, 88, 252, 216, 240, 11, 249, 139, 207, 191, 174, 121, 169, 243, 114, 239, 171, 169, 175, 58, 199, 35, 199, 31, 188, 206, 121, 61, 241, 166, 252, 173, 206, 219, 125, 239, 184, 239, 186, 223, 199, 189, 31, 153, 40, 252, 64, 254, 80, 243, 209, 250, 99, 199, 167, 208, 79, 247, 62, 231, 124, 254, 252, 47, 247, 132, 243, 251, 37, 210, 159, 51, 0, 0, 0, 4, 103, 65, 77, 65, 0, 0, 177, 142, 124, 251, 81, 147, 0, 0, 0, 32, 99, 72, 82, 77, 0, 0, 122, 37, 0, 0, 128, 131, 0, 0, 249, 255, 0, 0, 128, 233, 0, 0, 117, 48, 0, 0, 234, 96, 0, 0, 58, 152, 0, 0, 23, 111, 146, 95, 197, 70, 0, 0, 0, 138, 73, 68, 65, 84, 120, 218, 148, 210, 59, 14, 1, 81, 24, 5, 224, 111, 24, 180, 30, 209, 41, 232, 109, 200, 2, 148, 86, 33, 182, 160, 83, 171, 68, 236, 69, 45, 81, 88, 129, 66, 66, 113, 53, 183, 24, 220, 25, 185, 39, 57, 213, 127, 206, 255, 46, 66, 8, 114, 208, 146, 137, 108, 67, 129, 37, 230, 232, 226, 85, 163, 107, 71, 94, 74, 204, 176, 192, 184, 193, 208, 193, 29, 123, 24, 96, 133, 7, 66, 13, 159, 81, 211, 175, 102, 89, 55, 24, 54, 177, 253, 15, 244, 176, 77, 136, 119, 113, 190, 36, 70, 56, 84, 196, 39, 12, 255, 109, 110, 130, 115, 228, 244, 59, 88, 38, 12, 55, 28, 99, 133, 235, 207, 29, 114, 95, 227, 61, 0, 251, 29, 41, 167, 16, 244, 38, 66, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
$PNGStream = New-Object IO.MemoryStream($PNGImage, 0, $PNGImage.Length)
[System.Drawing.Image] $theImage = [System.Drawing.Image]::FromStream($PNGStream)


#Remove User button ============================================================
$deleteUserButton = New-Object System.Windows.Forms.Button
$deleteUserButton.Location = New-Object System.Drawing.Size(470,170)
$deleteUserButton.Size = New-Object System.Drawing.Size(120,23)
$deleteUserButton.Text = "Remove User(s)"
$deleteUserButton.Image = $theImage
$deleteUserButton.TextImageRelation = [System.Windows.Forms.TextImageRelation]::TextBeforeImage
$deleteUserButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top


$deleteUserButton.Add_Click(
{
	$StatusLabel.Text = ""
	$addUserButton.enabled = $false
	$deleteUserButton.enabled = $false
	$addOrbitButton.enabled = $false
	$deleteOrbitButton.enabled = $false
	$refreshButton.enabled = $false
	$FilterButton.enabled = $false
	$FindUserButton.enabled = $false
	$exportSCVButton.Enabled = $false
	
	$StatusLabel.Text = "Please Wait: Removing User(s)"
	[System.Windows.Forms.Application]::DoEvents()
	Write-Host ""
	Write-Host "Removing user from group..." -foreground "yellow"
	Write-Host ""
	if($Script:SkypeForBusinessAvailable)
	{
		RemoveUserFromGroupSkype
	}
	else
	{
		RemoveUserFromGroup
		Start-Sleep -s 1
	}
	#THIS CODE CAN BE USED TO REFRESH CONFIGURATION FROM THE SYSTEM AFTER EVERY SETTING CHANGE. HOWEVER, THIS CAN SLOW THINGS DOWN. SO IT'S NOT DONE ANYMORE.
	<#
	if($Script:SkypeForBusinessAvailable)
	{
		LoadGroupsSkype
	}
	else
	{
		LoadGroups
	}
	#>
	#Update Group List
	UpdateGroupsList
		
	$StatusLabel.Text = ""
	UpdateUsersList
	
	$addUserButton.enabled = $true
	$deleteUserButton.enabled = $true
	$addOrbitButton.enabled = $true
	$deleteOrbitButton.enabled = $true
	$refreshButton.enabled = $true
	$FilterButton.enabled = $true
	$FindUserButton.enabled = $true
	$exportSCVButton.Enabled = $true
})
$objForm.Controls.Add($deleteUserButton)



#Orbit Start Text box ============================================================
$orbitStartTextBox= new-object System.Windows.Forms.textbox
$orbitStartTextBox.location = new-object system.drawing.size(70,390)
$orbitStartTextBox.size= new-object system.drawing.size(100,15)
$orbitStartTextBox.text = ""   
$orbitStartTextBox.tabIndex = 2
$orbitStartTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$objform.controls.add($orbitStartTextBox)

#Start Label ============================================================
$objStartLabel = New-Object System.Windows.Forms.Label
$objStartLabel.Location = New-Object System.Drawing.Size(5,395) 
$objStartLabel.Size = New-Object System.Drawing.Size(70,15) 
$objStartLabel.Text = "Range Start: "
$objStartLabel.TabStop = $false
$objStartLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($objStartLabel)


#Orbit Stop Text box ============================================================
$orbitStopTextBox= new-object System.Windows.Forms.textbox
$orbitStopTextBox.location = new-object system.drawing.size(70,410)
$orbitStopTextBox.size= new-object system.drawing.size(100,15)
$orbitStopTextBox.text = ""
$orbitStopTextBox.tabIndex = 3  
$orbitStopTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$objform.controls.add($orbitStopTextBox)

#Stop Label ============================================================
$objStopLabel = New-Object System.Windows.Forms.Label
$objStopLabel.Location = New-Object System.Drawing.Size(5,415) 
$objStopLabel.Size = New-Object System.Drawing.Size(70,15) 
$objStopLabel.Text = "Range End: "
$objStopLabel.TabStop = $false
$objStopLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($objStopLabel)


#Add Orbit button ============================================================
$addOrbitButton = New-Object System.Windows.Forms.Button
$addOrbitButton.Location = New-Object System.Drawing.Size(20,330)
$addOrbitButton.Size = New-Object System.Drawing.Size(150,23)
$addOrbitButton.Text = "Add / Edit Orbit"
$addOrbitButton.TabIndex = 5
$addOrbitButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$addOrbitButton.Add_Click(
{
	$StatusLabel.Text = ""
	$orbitPool = $poolOrbitDropDownBox.SelectedItem.ToString()
	$name = $orbitNameTextBox.text
	$start = $orbitStartTextBox.text
	$stop = $orbitStopTextBox.text
	
	#Check if orbit already exists
	$OrbitExists = $false
	foreach($OrbitName in $objOrbitsListbox.Items)
	{
		#Write-Host "$orbitName EQUALS $name"
		if($orbitName -eq $name)
		{
			$OrbitExists = $true
			Write-Host "INFO: Orbit Exists. Changing Orbit." -foreground "Yellow"
			break
		}
		
	}
	if($OrbitExists)
	{
		$CurrentOrbitDetails = Invoke-Expression "Get-CsCallParkOrbit -Identity `"$name`" -Type GroupPickup"
		[string] $CurrentStart = $CurrentOrbitDetails.NumberRangeStart
		[string] $CurrentEnd = $CurrentOrbitDetails.NumberRangeEnd
		
		
		$regex1 = [regex] "^[\*|#]?[1-9]\d{0,7}$"
		$regex2 = [regex] "^[1-9][0-9]{0,8}$"
		[string] $regexStartResult = $regex1.match($CurrentStart)
		[string] $regexStartResult2 = $regex2.match($CurrentStart)
		[string] $regexEndResult = $regex1.match($CurrentEnd)
		[string] $regexEndResult2 = $regex2.match($CurrentEnd)
		
		
		if($regexStartResult -ne "" -or $regexStartResult2 -ne "" -or $regexEndResult -ne "" -or $regexEndResult2 -ne "")
		{
			$CurrentStart = $CurrentStart.Replace("#","").Replace("*","")
			$CurrentEnd = $CurrentEnd.Replace("#","").Replace("*","")
			
			$AttemptedStart = $start.Replace("#","").Replace("*","")
			$AttemptedEnd = $stop.Replace("#","").Replace("*","")
			
			[int]$AttemptStartNum = [convert]::ToInt32($AttemptedStart, 10)
			[int]$AttemptEndNum = [convert]::ToInt32($AttemptedEnd, 10)
			[int]$CurrentStartNum = [convert]::ToInt32($CurrentStart, 10)
			[int]$CurrentEndNum = [convert]::ToInt32($CurrentEnd, 10)
			
			if($AttemptStartNum -gt $CurrentStartNum -or $AttemptEndNum -lt $CurrentEndNum)
			{
				$a = new-object -comobject wscript.shell 
				$intAnswer = $a.popup("The range that you have selected is less than the previously configured range, any groups that contain users and fall outside of the new range will be retained but you will no longer be able to see them in this tool. Do you want to continue to edit the Orbit Range?",0,"Information",4) 
				if ($intAnswer -eq 6) 
				{
				
					$StatusLabel.Text = "Status: Editing Orbit..."
					$command = "Set-CsCallParkOrbit -Identity `"$name`" -Type GroupPickup -NumberRangeStart `"$start`" -NumberRangeEnd `"$stop`" -CallParkService `"$orbitPool`""
					Write-Host "COMMAND: $command" -foreground "green"
					Invoke-Expression $command
					$StatusLabel.Text = ""

				}else
				{Write-Host "INFO: Cancelled." -foreground "Yellow"}
			}
			else
			{
				$StatusLabel.Text = "Status: Editing Orbit..."
				$command = "Set-CsCallParkOrbit -Identity `"$name`" -Type GroupPickup -NumberRangeStart `"$start`" -NumberRangeEnd `"$stop`" -CallParkService `"$orbitPool`""
				Write-Host "COMMAND: $command" -foreground "green"
				Invoke-Expression $command
				$StatusLabel.Text = ""
			}	
		}
		else
		{
			Write-Host "ERROR: Range Start or End values are not allowed." -foreground "red"
		}	
	}
	else
	{
		$command = "New-CsCallParkOrbit -Identity `"$name`" -Type GroupPickup -NumberRangeStart `"$start`" -NumberRangeEnd `"$stop`" -CallParkService `"$orbitPool`""
		Write-Host "COMMAND: $command" -foreground "green"
		Invoke-Expression $command
	}
	UpdateOrbitsList
	UpdateGroupsList
})
$objForm.Controls.Add($addOrbitButton)


#Orbit Name Text box ============================================================
$orbitNameTextBox = new-object System.Windows.Forms.textbox
$orbitNameTextBox.location = new-object system.drawing.size(70,360)
$orbitNameTextBox.size = new-object system.drawing.size(100,15)
$orbitNameTextBox.text = ""   
$orbitNameTextBox.tabIndex = 1
$orbitNameTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$objform.controls.add($orbitNameTextBox)

#Name Label ============================================================
$objNameLabel = New-Object System.Windows.Forms.Label
$objNameLabel.Location = New-Object System.Drawing.Size(25,365) 
$objNameLabel.Size = New-Object System.Drawing.Size(50,15) 
$objNameLabel.Text = "Name: "
$objNameLabel.TabStop = $false
$objNameLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($objNameLabel)


#Delete Orbit button ============================================================
$deleteOrbitButton = New-Object System.Windows.Forms.Button
$deleteOrbitButton.Location = New-Object System.Drawing.Size(20,460)
$deleteOrbitButton.Size = New-Object System.Drawing.Size(150,23)
$deleteOrbitButton.Text = "Delete Orbit"
$deleteOrbitButton.TabStop = $false
$deleteOrbitButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$deleteOrbitButton.Add_Click(
{
	$StatusLabel = ""
	$orbitName = $objOrbitsListbox.SelectedItem
	if($orbitName -ne $null)
	{
		$a = new-object -comobject wscript.shell 
		$intAnswer = $a.popup("Are you sure you want to delete this entire Orbit Range?",0,"Information",4) 
		if ($intAnswer -eq 6) 
		{
		
			$StatusLabel.Text = "Status: Deleting Orbit..."
			DeleteOrbit
			UpdateOrbitsList
			UpdateGroupsList
			$StatusLabel.Text = ""

		}else
		{Write-Host "INFO: Cancelled." -foreground "Yellow"}
		
	}
})
$objForm.Controls.Add($deleteOrbitButton)


# Add Orbits Dropdown box ============================================================
$poolOrbitDropDownBox = New-Object System.Windows.Forms.ComboBox 
$poolOrbitDropDownBox.Location = New-Object System.Drawing.Size(20,435) 
$poolOrbitDropDownBox.Size = New-Object System.Drawing.Size(150,20) 
$poolOrbitDropDownBox.DropDownHeight = 60 
$poolOrbitDropDownBox.tabIndex = 4
$poolOrbitDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($poolOrbitDropDownBox) 

Get-CSService -UserServer | where-object {$_.version -eq "6" -or $_.version -eq "7" -or $_.version -eq "8"} | select-object PoolFQDN | ForEach-Object {[void] $poolOrbitDropDownBox.Items.Add($_.PoolFQDN)}

$numberOfItems = $poolOrbitDropDownBox.Items.count
if($numberOfItems -gt 0)
{
	$poolOrbitDropDownBox.SelectedIndex = 0
}


#Refresh button ============================================================
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Size(195,457)
$refreshButton.Size = New-Object System.Drawing.Size(150,20)
$refreshButton.Text = "Refresh"
$refreshButton.TabIndex = 7
$refreshButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$refreshButton.Add_Click(
{
	$StatusLabel.Text = "Refreshing configuration from system..."
	$addUserButton.enabled = $false
	$deleteUserButton.enabled = $false
	$addOrbitButton.enabled = $false
	$deleteOrbitButton.enabled = $false
	$refreshButton.enabled = $false
	$FilterButton.enabled = $false
	$FindUserButton.enabled = $false
	$exportSCVButton.Enabled = $false
	[System.Windows.Forms.Application]::DoEvents()
	
	RefreshInterface
	
	$StatusLabel.Text = ""
	
	$addUserButton.enabled = $true
	$deleteUserButton.enabled = $true
	$addOrbitButton.enabled = $true
	$deleteOrbitButton.enabled = $true
	$refreshButton.enabled = $true
	$FilterButton.enabled = $true
	$FindUserButton.enabled = $true
	$exportSCVButton.Enabled = $true

	
})
$objForm.Controls.Add($refreshButton)


#Refresh button ============================================================
$exportSCVButton = New-Object System.Windows.Forms.Button
$exportSCVButton.Location = New-Object System.Drawing.Size(195,482)
$exportSCVButton.Size = New-Object System.Drawing.Size(150,20)
$exportSCVButton.Text = "Export CSV"
$exportSCVButton.TabIndex = 8
$exportSCVButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$exportSCVButton.Add_Click(
{
	$StatusLabel.Text = "Exporting CSV..."
	$addUserButton.enabled = $false
	$deleteUserButton.enabled = $false
	$addOrbitButton.enabled = $false
	$deleteOrbitButton.enabled = $false
	$refreshButton.enabled = $false
	$FilterButton.enabled = $false
	$FindUserButton.enabled = $false
	$exportSCVButton.Enabled = $false
	[System.Windows.Forms.Application]::DoEvents()
	
	$filename = ""
	
	Write-Host "INFO: Exporting..." -foreground "yellow"
	[string] $pathVar = "c:\"
	$Filter="All Files (*.*)|*.*"
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$objDialog = New-Object System.Windows.Forms.SaveFileDialog
	#$objDialog.InitialDirectory = 
	$objDialog.FileName = "GroupCallPickup.csv"
	$objDialog.Filter = $Filter
	$objDialog.Title = "Export File Name"
	$objDialog.CheckFileExists = $false
	$Show = $objDialog.ShowDialog()
	if ($Show -eq "OK")
	{
		[string] $filename = $objDialog.FileName
	}
	
	Write-Host "INFO: $filename" -foreground "yellow"
	if($filename -ne "")
	{
	
		$csv = "`"CallPickupGroup`",`"User`"`r`n"
		
		$objGroupsListboxItemSelected = $objGroupsListbox.SelectedItems[0].Text
		$objUsersLabel.Text = "Users in Group $objGroupsListboxItemSelected"
		
		foreach($groupItem in $objGroupsListbox.items)
		{
			$GroupsListboxItem = $groupItem.Text
			foreach($group in $groups)
			{
				[string] $GroupNumber = $group.GroupNumber
				
				if($GroupNumber -eq $GroupsListboxItem)
				{
						$CallPickupGroup = $group.GroupNumber
						$User = $group.UserSipAddress
								
						$csv += "`"" +$CallPickupGroup+"`",`"" +$User+"`"`r`n"
				}
			}
		}
		
		#Excel seems to only like UTF-8 for CSV files...
		$csv | out-file -Encoding UTF8 -FilePath $filename -Force
		Write-Host "Completed Export." -foreground "yellow"
		
	}
	else
	{
		Write-Host "INFO: No filename selected." -foreground "Yellow"
	}
	
	$StatusLabel.Text = ""
	
	$addUserButton.enabled = $true
	$deleteUserButton.enabled = $true
	$addOrbitButton.enabled = $true
	$deleteOrbitButton.enabled = $true
	$refreshButton.enabled = $true
	$FilterButton.enabled = $true
	$FindUserButton.enabled = $true
	$exportSCVButton.Enabled = $true

	
})
$objForm.Controls.Add($exportSCVButton)


#CHANGED ALL THE GROUP OPERATION IN VERSION 2.0 THESE GUI ELEMENTS ARE NOT NEEDED ANYMORE.
<#
#Add Group button ============================================================
$addGroupButton = New-Object System.Windows.Forms.Button
$addGroupButton.Location = New-Object System.Drawing.Size(195,330)
$addGroupButton.Size = New-Object System.Drawing.Size(150,23)
$addGroupButton.Text = "Add Group"
$addGroupButton.TabIndex = 7
$addGroupButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$addGroupButton.Add_Click(
{
	$StatusLabel.Text = ""
	AddGroup
})
$objForm.Controls.Add($addGroupButton)


#Group Number Text box ============================================================
$groupNumberTextBox = new-object System.Windows.Forms.textbox
$groupNumberTextBox.location = new-object system.drawing.size(255,360)
$groupNumberTextBox.size= new-object system.drawing.size(90,15)
$groupNumberTextBox.text = ""
$groupNumberTextBox.TabIndex = 6
$groupNumberTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left

$groupNumberTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		AddGroup
	}
})
$objform.controls.add($groupNumberTextBox)

#Group Number Label ============================================================
$groupNumberLabel = New-Object System.Windows.Forms.Label
$groupNumberLabel.Location = New-Object System.Drawing.Size(195,360) 
$groupNumberLabel.Size = New-Object System.Drawing.Size(60,15) 
$groupNumberLabel.Text = "Group No: "
$groupNumberLabel.TabStop = $false
$groupNumberLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($groupNumberLabel)
#>

<#
# Add the Close button ============================================================
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(475,483)
$CancelButton.Size = New-Object System.Drawing.Size(100,20)
$CancelButton.Text = "Close"
$CancelButton.Add_Click({$objForm.Close()})
$CancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$objForm.Controls.Add($CancelButton)
#>

# Add the Status Label ============================================================
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Location = New-Object System.Drawing.Size(20,503) 
$StatusLabel.Size = New-Object System.Drawing.Size(420,15) 
$StatusLabel.Text = ""
$StatusLabel.forecolor = "red"
$StatusLabel.TabStop = $false
$StatusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($StatusLabel)


$MyLinkLabel = New-Object System.Windows.Forms.LinkLabel
$MyLinkLabel.Location = New-Object System.Drawing.Size(470,3)
$MyLinkLabel.Size = New-Object System.Drawing.Size(170,15)
$MyLinkLabel.DisabledLinkColor = [System.Drawing.Color]::Red
$MyLinkLabel.VisitedLinkColor = [System.Drawing.Color]::Blue
$MyLinkLabel.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$MyLinkLabel.LinkColor = [System.Drawing.Color]::Navy
$MyLinkLabel.TabStop = $False
$MyLinkLabel.Text = "www.myskypelab.com"
$MyLinkLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$MyLinkLabel.add_click(
{
	 [system.Diagnostics.Process]::start("http://www.myskypelab.com")
})
$objForm.Controls.Add($MyLinkLabel)


$ToolTip = New-Object System.Windows.Forms.ToolTip 
$ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow 
$ToolTip.IsBalloon = $true 
$ToolTip.InitialDelay = 2000 
$ToolTip.ReshowDelay = 1000 
$ToolTip.AutoPopDelay = 10000
#$ToolTip.ToolTipTitle = "Help:"
$ToolTip.SetToolTip($addUserButton, "This button will add a user from the list below to the selected Pickup Group.") 
$ToolTip.SetToolTip($deleteUserButton, "This button will remove the selected user from the Pickup Group") 
$ToolTip.SetToolTip($addOrbitButton, "This button will add a new Orbit Range to the system.") 
$ToolTip.SetToolTip($deleteOrbitButton, "This button will delete an Orbit Range from the system.") 
$ToolTip.SetToolTip($FilterButton, "This button will filter the user list based on the string in the text box.") 
$ToolTip.SetToolTip($refreshButton, "This button will refresh all the information from the system.`r`nThis can be handy if config changes have been done since the tool was opened.") 


# Update the Orbit fields ============================================================
function UpdateSelectedOrbit
{
	$orbitName = $objOrbitsListbox.SelectedItem
	if($orbitName -ne $null)
	{
	$orbitDetails = Invoke-Expression "Get-CsCallParkOrbit -Identity `"$orbitName`""
	$orbitStartTextBox.text = $orbitDetails.NumberRangeStart
	$orbitStopTextBox.text = $orbitDetails.NumberRangeEnd
	$orbitNameTextBox.text = $orbitDetails.Identity
	$loopIndex = 0
	foreach($item in $orbitPoolComboBox)
	{
		if($item -eq $orbitDetails.CallParkServerFqdn)
		{
			$orbitPoolComboBox.SelectedIndex=$loopIndex
			break
		}
		$loopIndex++
	}
	}
}

# Update the Orbits listbox ============================================================
function UpdateOrbitsList
{
	$objOrbitsListbox.items.Clear()
	# Add Lync Orbits
	Get-CsCallParkOrbit -type GroupPickup | select-object Identity | ForEach-Object {[void] $objOrbitsListbox.Items.Add($_.Identity)}
	
	#Fill the Groups List...
	$allOrbits = Get-CsCallParkOrbit -type GroupPickup
	
	Write-Host "INFO: Updating Groups List from Orbits..." -foreground "Yellow"
	Write-Host
	
	$objGroupsListbox.items.Clear()
	foreach($orbit in $allOrbits)
	{
		[string]$startNumber = $orbit.NumberRangeStart
		[string]$endNumber = $orbit.NumberRangeEnd
		
		Write-Host
		Write-Host "Orbit Range: $startNumber - $endNumber" -foreground "green"
		Write-Host
		
		$containsHash = $false
		$containsStar = $false
		
		if($startNumber.Contains("#"))
		{
			$startNumberCropped = $startNumber.Replace("#","")
			$endNumberCropped = $endNumber.Replace("#","")

			[int]$startNum = [convert]::ToInt32($startNumberCropped, 10)
			[int]$endNum = [convert]::ToInt32($endNumberCropped, 10)
			
			$containsHash = $true
			
		}
		elseif($startNumber.Contains("*"))
		{
			$startNumberCropped = $startNumber.Replace("*","")
			$endNumberCropped = $endNumber.Replace("*","")
			
			[int]$startNum = [convert]::ToInt32($startNumberCropped, 10)
			[int]$endNum = [convert]::ToInt32($endNumberCropped, 10)			
			
			$containsStar = $true
		}
		else
		{
			[int]$startNum = [convert]::ToInt32($startNumber, 10)
			[int]$endNum = [convert]::ToInt32($endNumber, 10)
		}
		
		$break = $true		
		$currentNum = $startNum
		while($break)
		{
			if($currentNum -eq $endNum)
			{
				[string]$currentString = [convert]::ToString($currentNum)
				if($containsHash)
				{
					$currentString = "#${currentString}"
				}
				elseif($containsStar)
				{
					$currentString = "*${currentString}"
				}
				Write-Host "Adding Group: $currentString" -foreground "green"
				[void] $objGroupsListbox.Items.add($currentString)
				$break = $false
			}
			else
			{
				[string]$currentString = [convert]::ToString($currentNum)
				if($containsHash)
				{
					$currentString = "#${currentString}"
				}
				elseif($containsStar)
				{
					$currentString = "*${currentString}"
				}
				Write-Host "Adding Group: $currentString" -foreground "green"
				[void] $objGroupsListbox.Items.add($currentString)
			}
			$currentNum++
		}
	}
	Write-Host
}

# Delete and Orbit ============================================================
function DeleteOrbit
{
	$deleteOrbit = $objOrbitsListbox.SelectedItem
	$command = "Remove-CsCallParkOrbit -Identity `"$deleteOrbit`""
	Write-Host "COMMAND: $command" -foreground "green"
	Invoke-Expression $command
}

# Add Group ============================================================
# DON'T NEED THIS FUNCTION WITH THE NEW 2.0 OPERATION. LEAVE IT JUST IN CASE...
function AddGroup
{
	$groupNumber = $groupNumberTextBox.text
	
	$MatchStringArray = @()
	$MatchLengthArray = @()
	foreach($Orbit in $objOrbitsListbox.items)
	{
		$GroupInfo = Get-CsCallParkOrbit -Identity "$Orbit"
		#NumberRangeStart   : #100
		#NumberRangeEnd     : #101

		[string]$StartString = $GroupInfo.NumberRangeStart
		[string]$EndString = $GroupInfo.NumberRangeEnd
		
		$loopNo = 0
		[string]$MatchString = ""
		while($StartString.length -gt $loopNo)
		{
			if($EndString[$loopNo] -eq $StartString[$loopNo])
			{
				[string]$MatchString += $StartString[$loopNo]
			}
			else
			{
				$MatchLengthArray += $StartString.length
				
				#Capture the end bit and make the Regex
				$EndBit1 = ""
				$EndBit2 = ""
				while($StartString.length -gt $loopNo)
				{
					$EndBit1 += $StartString[$loopNo]
					$EndBit2 += $EndString[$loopNo]
					$loopNo++
				}
				$MatchString = "$MatchString[$EndBit1-$EndBit2]"
				$MatchStringArray += $MatchString
				#Write-Host "FINAL MATCH STRING: $MatchString" #DEBUGGING
 			}
			$loopNo++
		}
	}
	
	$MatchedOrbit = $false
	
	$loopNo = 0
	foreach($MatchString in $MatchStringArray)
	{
		if($MatchString -ne "")
		{
			if($MatchLengthArray[$loopNo] -eq $groupNumber.length)
			{
				$MatchString = $MatchString.Replace("*", "\*")
				
				if($groupNumber -imatch $MatchString)
				{
					$MatchedOrbit = $true
				}
			}
			else
			{
				Write-Host "ERROR: The group number is not the same length as the Orbit range"
			}
		}
		else
		{
			Write-Host "Match string has no value..."
		}
		$loopNo++
	}
	
	if($MatchedOrbit)
	{
		$regex1 = [regex] "^[\*|#]?[1-9]\d{0,7}$"
		$regex2 = [regex] "^[1-9][0-9]{0,8}$"
		[string] $regexResult = $regex1.match($groupNumber)
		[string] $regexResult2 = $regex2.match($groupNumber)
		
		#Check if group is in the list already
		$groupAlreadyExists = $false
		foreach($item in $objGroupsListbox.Items)
		{
			[string] $listGroup = $item
			if($groupNumber -eq $listGroup)
			{
				$groupAlreadyExists = $true
				Write-Host "ERROR: Group already exists..." -foreground "Red"
			}
		}
		
		if($groupNumber -ne "" -and $groupAlreadyExists -eq $false)
		{
			if($regexResult -ne "" -or $regexResult2 -ne "")
			{
				#Write-Host "Information: Group Number format OK." -foreground "yellow"
				#Write-Host ""
				[void]$objGroupsListbox.Items.Add($groupNumber)
			}
			else
			{
				Write-Host "ERROR: Group number is not in the range ([*|#]?[1-9]\d{0,7})|([1-9]\d{0,8})" -foreground "red"
				Write-Host ""
			}
		}
		$listboxIndex = 0
		foreach($item in $objGroupsListbox.Items)
		{
			[string] $listboxItem = $item
			[string] $groupNumberString = $groupNumber
			if($listboxItem -eq $groupNumberString)
			{
				$objGroupsListbox.SelectedIndex = $listboxIndex
				break
			}
			$listboxIndex++
		}
		UpdateUsersList
		$itemsInUsersList = $objUsersListbox.Items.Count
		if($itemsInUsersList -gt 0)
		{$objUsersListbox.SelectedIndex = 0}
	}
	else
	{
		Write-Host "ERROR: The group number entered does not match an existing Orbit Range. Please create the Orbit Range first before trying to create the group." -foreground "red"
		$StatusLabel.Text = "ERROR: The group number entered does not match an existing Orbit Range."
	}
}

# Add a User to a Group ============================================================
function AddUserToGroup
{
	$ErrorOccured = $false
	[string] $pickupGroup =  $objGroupsListbox.SelectedItems[0].Text #Currently Selected Group
	foreach ($objItem in $objLyncUsersListbox.SelectedItems)
	{
	
		Write-Host "-------------------------------------------------------------"
		Write-Host "ADDING USER" -foreground "Green"
		[string] $userSipAddress = $objItem
		#$pickupPool = $poolOrbitDropDownBox.SelectedItem.ToString()
		$homePool = Get-CsUser -identity "sip:$userSipAddress" -erroraction 'silentlycontinue'
		if($homePool -eq $null)
		{
			$homePool = Get-CsCommonAreaPhone -identity "sip:$userSipAddress" -erroraction 'silentlycontinue'
			if($homePool -ne $null)
			{
				#write-host "Found Common Area Device: $homePool" -foreground yellow
			}
		}
		[string] $pickupPool = $homePool.RegistrarPool		
			
		$command = "$SEFAUTILPath2013 $userSipAddress /Server:$pickupPool /enablegrouppickup:$pickupGroup /verbose"
		Write-Host "COMMAND: $command" -foreground "green"
		
		if (Test-Path $SEFAUTILPath2013)
		{
			#Launch SEFAUTIL
			$pinfo = New-Object System.Diagnostics.ProcessStartInfo
			$pinfo.FileName = $SEFAUTILPath2013
			$pinfo.RedirectStandardError = $true
			$pinfo.RedirectStandardOutput = $true
			$pinfo.UseShellExecute = $false
			$pinfo.Arguments = "$userSipAddress /Server:$pickupPool /enablegrouppickup:$pickupGroup /verbose"
			$p = New-Object System.Diagnostics.Process
			$p.StartInfo = $pinfo
			$p.Start() | Out-Null
			$p.WaitForExit()
			[string]$stdout = $p.StandardOutput.ReadToEnd()
			[string]$stderr = $p.StandardError.ReadToEnd()
			
			if($stderr -ne "")
			{
				$ErrorOccured = $true
				Write-Host "ERROR Reported from SEFAUtil:" -Foreground "red"
				Write-Host "$stderr" -Foreground "red"
			}
			elseif($stdout -eq "")
			{
				$ErrorOccured = $true
				Write-Host "ERROR: No Response from SEFAUTIL. Check that SEFAUTIL is installed correctly." -foreground "red"
			}
			elseif($stdout -ne "")
			{
				Write-Host "INFO: The following was reported by SEFAUtil:" -foreground "yellow"
				Write-Host "$stdout"
				
				$regexSafeGroup = $pickupGroup.Replace("*","\*")
				if($stdout -imatch "Group Pickup Orbit: sip:${regexSafeGroup}")
				{
					foreach($group in $groups)
					{
						if($group.UserSipAddress -eq $userSipAddress)
						{
							Write-Host "Removing user from existing Group..." -foreground "Green"
							RemoveUserFromGroupArray -sipAddress "${userSipAddress}"
							break
						}
					}
					Write-Host "Adding User to new Group..." -foreground "Green"
					AddUserToGroupArray -SipAddress $userSipAddress -CPGNumber $pickupGroup -UserPoolIdentity $pickupPool
				}
			}
			
			Write-Host "-------------------------------------------------------------"
	
		}
		else
		{
			write-host ""
			write-host "Error: SEFAUTIL not found in location. Please install Lync 2013 reskit tools." -foreground "red"
			write-host ""
		}
	}
	if($ErrorOccured)
	{
		RefreshInterface #Not sure what happened so refresh interface
		#Reset the Group
		$item = $objGroupsListbox.FindItemWithText($pickupGroup)
		$item.Selected = $true
		$item.Focused = $true
		UpdateUsersList
	}	
	$itemsInUserList = $objUsersListbox.Items.Count
	if($itemsInUserList -gt 0)
	{$objUsersListbox.SelectedIndex = 0}
}

# Remove a User from a Group ============================================================
function RemoveUserFromGroup
{
	$ErrorOccured = $false
	[string] $pickupGroup =  $objGroupsListbox.SelectedItems[0].Text #Currently Selected Group
	foreach ($objItem in $objUsersListbox.SelectedItems)
	{
		Write-Host "-------------------------------------------------------------"
		Write-Host "REMOVING USER" -foreground "Green"
		$userSipAddress = $objItem

		$homePool = Get-CsUser -identity "sip:$userSipAddress" -erroraction 'silentlycontinue'
		if($homePool -eq $null)
		{
			$homePool = Get-CsCommonAreaPhone -identity "sip:$userSipAddress" -erroraction 'silentlycontinue'
			if($homePool -ne $null)
			{
				#write-host "Found Common Area Device: $homePool" -foreground yellow
			}
		}
		[string] $pickupPool = $homePool.RegistrarPool	
		
		$command = "$SEFAUTILPath2013 $userSipAddress /Server:$pickupPool /disablegrouppickup /verbose"
		Write-Host "COMMAND: $command" -foreground "green"
		
		if (Test-Path $SEFAUTILPath2013)
		{
			#Launch SEFAUTIL
			$pinfo = New-Object System.Diagnostics.ProcessStartInfo
			$pinfo.FileName = $SEFAUTILPath2013
			$pinfo.RedirectStandardError = $true
			$pinfo.RedirectStandardOutput = $true
			$pinfo.UseShellExecute = $false
			$pinfo.Arguments = "$userSipAddress /Server:$pickupPool /disablegrouppickup /verbose"
			$p = New-Object System.Diagnostics.Process
			$p.StartInfo = $pinfo
			$p.Start() | Out-Null
			$p.WaitForExit()
			[string]$stdout = $p.StandardOutput.ReadToEnd()
			[string]$stderr = $p.StandardError.ReadToEnd()

			if($stderr -ne "")
			{
				$ErrorOccured = $true
				Write-Host "ERROR Reported from SEFAUtil:" -Foreground "red"
				Write-Host "$stderr" -Foreground "red"
			}
			elseif($stdout -eq "")
			{
				$ErrorOccured = $true
				Write-Host "ERROR: No Response from SEFAUTIL. Check that SEFAUTIL is installed correctly." -foreground "red"
			}
			elseif($stdout -ne "")
			{
				Write-Host "INFO: The following was reported by SEFAUtil:" -foreground "yellow"
				Write-Host "$stdout"
				
				if(!($stdout -imatch "Group Pickup Orbit"))
				{
					#Version 2.0 needs to update the groups itself
					Write-Host "Remove Success" -foreground "Green"
					RemoveUserFromGroupArray -sipAddress "${userSipAddress}"
				}
			}
			
			Write-Host "-------------------------------------------------------------"
		
		}
		else
		{
			write-host ""
			write-host "Error: SEFAUTIL not found in location. Please install Lync 2013 reskit tools." -foreground "red"
			write-host ""
		}	
	}
	if($ErrorOccured)
	{
		RefreshInterface #Not sure what happened so refresh interface
		$item = $objGroupsListbox.FindItemWithText($pickupGroup)
		$item.Selected = $true
		$item.Focused = $true
		UpdateUsersList
	}
}

# Add a User from a Skype for Buuiness Group ============================================================
function AddUserToGroupSkype
{
	$ErrorOccured = $false
	[string] $pickupGroup =  $objGroupsListbox.SelectedItems[0].Text #Currently Selected Group
	foreach ($objItem in $objLyncUsersListbox.SelectedItems)
	{
		Write-Host "-------------------------------------------------------------"
		Write-Host "ADDING USER" -foreground "Green"
		$userSipAddress = $objItem
		
		$UserGroup = Invoke-Expression "Get-CsGroupPickupUserOrbit -user sip:${userSipAddress} -verbose"
		[string]$CPGNumber = $UserGroup.Orbit
		
		if($CPGNumber -eq "" -or $CPGNumber -eq $null)
		{
			$command = "New-CsGroupPickupUserOrbit sip:${userSipAddress} -Orbit `"$pickupGroup`" -verbose"
			Write-Host
			Write-Host $command -foreground "green"
			Write-Host
			$response = Invoke-Expression $command
			
			
			if($response.Orbit -eq $pickupGroup)
			{
				Write-Host "INFO: New User Pickup Setting Success" -foreground "Green"
				$UserInfo = Invoke-Expression "Get-CsUser -identity `"sip:${userSipAddress}`""
				$RegistrarPool = $UserInfo.RegistrarPool
			
				AddUserToGroupArray -SipAddress ${userSipAddress} -CPGNumber $pickupGroup -UserPoolIdentity $RegistrarPool
			}
			else
			{
				$ErrorOccured = $true
				Write-Host "ERROR: Setting Failed" -foreground "Red"
				Write-Host "RESPONSE: $response" -foreground "Red"
			}
		}
		else
		{
			$command = "Set-CsGroupPickupUserOrbit sip:${userSipAddress} -Orbit `"$pickupGroup`" -verbose"
			Write-Host $command -foreground "green"
			$response = Invoke-Expression $command
			#Write-Host "RESPONSE: $response" #DEBUGGING
			
			if($response.Orbit -eq $pickupGroup)
			{
				Write-Host "Change User Pickup Setting Success" -foreground "Green"
				
				RemoveUserFromGroupArray -sipAddress "${userSipAddress}"
				
				$UserInfo = Invoke-Expression "Get-CsUser -identity `"sip:${userSipAddress}`""
				$RegistrarPool = $UserInfo.RegistrarPool
				
				AddUserToGroupArray -SipAddress ${userSipAddress} -CPGNumber $pickupGroup -UserPoolIdentity $RegistrarPool
			}
			else
			{
				$ErrorOccured = $true
				Write-Host "ERROR: Setting Failed" -foreground "Red"
				Write-Host "RESPONSE: $response" -foreground "Red"
			}
		}
		Write-Host "-------------------------------------------------------------"
	}
	if($ErrorOccured)
	{
		RefreshInterface #Not sure what happened due to error. So refresh interface.
		$item = $objGroupsListbox.FindItemWithText($pickupGroup)
		$item.Selected = $true
		$item.Focused = $true
		UpdateUsersList
	}
}

# Remove a User from a Skype for Business Group ============================================================
function RemoveUserFromGroupSkype
{
	$ErrorOccured = $false
	[string] $pickupGroup =  $objGroupsListbox.SelectedItems[0].Text #Currently Selected Group
	foreach ($objItem in $objUsersListbox.SelectedItems)
	{
		Write-Host "-------------------------------------------------------------"
		Write-Host "REMOVING USER" -foreground "Green"
		$userSipAddress = $objItem
				
		$command = "Remove-CsGroupPickupUserOrbit -user sip:${userSipAddress} -verbose"
		Write-Host $command -foreground "green"
				
		$ErrorVar = $null
		$response = Invoke-Expression $command -ErrorVariable ErrorVar
		
				
		if((!($ErrorVar.length -gt 0)) -and $response.Orbit -eq $null)
		{
			Write-Host "Remove Success" -foreground "Green"
			RemoveUserFromGroupArray -sipAddress "${userSipAddress}"
		}
		else
		{
			$ErrorOccured = $true
			Write-Host "ERROR: Remove Error." -foreground "Red"
			#Write-Host "$ErrorVar" -foreground "Red"
		}
		Write-Host "-------------------------------------------------------------"
	}
	if($ErrorOccured)
	{
		RefreshInterface #Not sure what happened so refresh interface
		$item = $objGroupsListbox.FindItemWithText($pickupGroup)
		$item.Selected = $true
		$item.Focused = $true
		UpdateUsersList
	}
}

function RemoveUserFromGroupArray([string] $SipAddress)
{
	Write-Host "INFO: Removing user from group array - $SipAddress" -foreground "Yellow"
	
	#Update the Groups Array with a new array that doesn't contain this SIP Address
	$TempArray = @()
	foreach($group in $groups)
	{
		if($group.UserSipAddress -ne $SipAddress)
		{
			$TempArray += @($group)
		}
	}
	
	$script:groups = $TempArray
}

function AddUserToGroupArray([string] $SipAddress, [string] $CPGNumber, [String] $UserPoolIdentity)
{
	Write-Host "INFO: Adding User to Group Array - $SipAddress $CPGNumber $UserPoolIdentity" -foreground "Yellow"
	$script:groups += @(@{"UserPool" = "$UserPoolIdentity";"GroupNumber" = "$CPGNumber";"UserSipAddress" = "$SipAddress";"Version" = 0})
}


# Load Groups Array ============================================================
function LoadGroups 
{
	
	$script:groups = @()
	write-host
	
	foreach($computer in $computers)
	{
		[string]$Server = $computer
		write-host "INFO: Getting user call pickup information from $Server" -foreground "Yellow"
		
		#Define SQL Connection String
		[string]$connstring = "server=$server\rtclocal;database=rtc;trusted_connection=true;"
	 
		#Define SQL Command
		[object]$command = New-Object System.Data.SqlClient.SqlCommand

		<#
		#OLD QUERY. THIS METHOD HAS ISSUES...
		# SQL query for Lync Server 2013
		$command.CommandText = "select distinct UserAtHost,convert(varchar(4000),convert(varbinary(4000),Data)) `
		from PublishedStaticInstance,Resource `
		where ResourceId = PublisherId `
		and convert(varchar(4000),convert(varbinary(4000),Data)) `
		like '%<list name=`"GroupPickupList`"><target uri=%'"
		#>
		
		#New 2.0 method
		$command.CommandText = "select distinct UserAtHost,Version,convert(varchar(4000),convert(varbinary(4000),Data)) `
		from PublishedStaticInstance,Resource `
		where ResourceId = PublisherId `
		and convert(varchar(4000),convert(varbinary(4000),Data)) `
		like '%<routing xmlns=%'"
		
		
		[object]$connection = New-Object System.Data.SqlClient.SqlConnection
		$connection.ConnectionString = $connstring
		try {
		$connection.Open()
		} catch [Exception] {
			write-host ""
			write-host "Lync 2013 Call Pickup Manager was unable to connect to database $server\rtclocal. Please check that the server is online. Also check that UDP 1434 and the Dynamic SQL TCP Port for the RTCLOCAL Named Instance are open in the Windows Firewall on $server." -foreground "red"
			write-host ""
			$StatusLabel.Text = "ERROR: Error connecting to $server. Refer to PowerShell window."
		}
		
		$command.Connection = $connection
		
	 
		[object]$sqladapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$sqladapter.SelectCommand = $command
	 
		[object]$results = New-Object System.Data.Dataset
		try {
		$recordcount = $sqladapter.Fill($results)
		} catch [Exception] {
			write-host ""
			write-host "Error running SQL on $server : $_" -foreground "red"
			write-host ""
		}
		$connection.Close()
		$data = $Results.Tables[0]

		
		foreach ($Row in $data)
		{ 
			
			[string]$callPickUpSetting = $Row[2]
			
			[string]$SipAddress = $Row[0]
			
			[string]$Version = [convert]::ToInt32($Row[1], 10)
			#Write-Host "$SipAddress - Version: $Version" #DEBUGGING
			
			$EditTheUser = $false
			$DidNotFindUser = $true
			$LoopNo = 0
			foreach($group in $groups)
			{
				if($group.UserSipAddress -eq $SipAddress)
				{
					#Write-Host "VERSION COMPARE $SipAddress : $Version gt " $group.Version #DEBUGGING
					if($Version -gt $group.Version)
					{
						#Write-Host "EDIT THE USER!" #DEBUGGING
						$EditTheUser = $true
						break
					}
					$DidNotFindUser = $false
				}
				$LoopNo++
			}
			
			if($EditTheUser -eq $true -or $DidNotFindUser -eq $true)
			{
				#Write-Host "CALL PICKUP SETTING: $callPickUpSetting" #DEBUGGING
				if($callPickUpSetting.Contains("<list name=`"GroupPickupList`">"))
				{
					if(!($callPickUpSetting -match "<list name=`"GroupPickupList`"></list>"))
					{
						$xmlSplitStart = $callPickUpSetting -split "<list name=`"GroupPickupList`">"
						$splitStart = $xmlSplitStart[1]
						$xmlSplitCallParkList = $splitStart -split "<//list>"
						$splitCallParkList = $xmlSplitCallParkList[0]
						
						if($splitCallParkList.Contains("target uri=`"sip:"))
						{
							#Write-Host "Found Target URI"
							$start = $splitCallParkList.IndexOf("<target uri=`"sip:")
							$end = $splitCallParkList.IndexOf(";phone-context=user-default")
							[string]$CPGNumber = $splitCallParkList.Substring($start+17,$end-$start-17)
							
											
							$UserPool = Get-CsPool | Where-Object {$_.Computers –contains $computer} | select-object Identity
							$UserPoolIdentity = $UserPool.Identity
							
							$GroupSipAddress = $SipAddress
							$homePool = Get-CsUser -identity "sip:${SipAddress}" -erroraction 'silentlycontinue'
							if($homePool -eq $null)
							{
								$homePool = Get-CsCommonAreaPhone -identity "sip:${SipAddress}" -erroraction 'silentlycontinue'
								if($homePool -ne $null)
								{
									#[string]$CommonSipAddress = $homePool.SipAddress
									#[string]$CommonDisplayName = $homePool.DisplayName
									#write-host "Found Common Area Device: $CommonSipAddress ($CommonDisplayName)" -foreground yellow
								}
							}
							[string] $homePoolRegistrar = $homePool.RegistrarPool
							
							
							if($homePoolRegistrar -eq $UserPoolIdentity -and $homePool.EnterpriseVoiceEnabled -eq $true)
							{
								#Write-Host "FOUND USER: $SipAddress $CPGNumber $UserPoolIdentity"   ##DEBUG
								if($EditTheUser)
								{
									#Write-Host "Editing Existing User: $SipAddress $CPGNumber $UserPoolIdentity $Version / " $script:groups[$LoopNo].UserSipAddress -foreground "yellow" #DEBUGGING
									$script:groups[$LoopNo].UserPool = $UserPoolIdentity
									$script:groups[$LoopNo].GroupNumber = $CPGNumber
									$script:groups[$LoopNo].Version = $Version
								}
								else
								{
									#Write-Host "Adding new user: $SipAddress $CPGNumber $UserPoolIdentity $Version"	-foreground "yellow" #DEBUGGING
									$script:groups += @(@{"UserPool" = "$UserPoolIdentity";"GroupNumber" = "$CPGNumber";"UserSipAddress" = "$SipAddress";"Version" = "$Version"}) 
								}
							}
						}
						else
						{
							#NO TARGET URI SPECIFIED IN DB ENTRY
							if($EditTheUser)
							{
								#Write-Host "Editing Existing User: $SipAddress $CPGNumber $UserPoolIdentity $Version / " $script:groups[$LoopNo].UserSipAddress -foreground "yellow" #DEBUGGING
								$TempArray = @()
								foreach($group in $groups)
								{
									if($group.UserSipAddress -ne $SipAddress)
									{
										$TempArray += @($group)
									}
								}
								
								$script:groups = $TempArray
								
							}
						}
					}
					else
					{
						#NO TARGET URI SPECIFIED IN DB ENTRY
						if($EditTheUser)
						{
							#Write-Host "Editing Existing User: $SipAddress $CPGNumber $UserPoolIdentity $Version / " $script:groups[$LoopNo].UserSipAddress -foreground "yellow" #DEBUGGING
							$TempArray = @()
							foreach($group in $groups)
							{
								if($group.UserSipAddress -ne $SipAddress)
								{
									$TempArray += @($group)
								}
							}
							$script:groups = $TempArray
						}
					}
				}
			}
			else
			{
				#NO GROUP CALL PICKUP SETTING FOUND IN DB ENTRY
				if($EditTheUser)
				{
					#Write-Host "Editing Existing User: $SipAddress $CPGNumber $UserPoolIdentity $Version / " $script:groups[$LoopNo].UserSipAddress -foreground "yellow" #DEBUGGING
					$TempArray = @()
					foreach($group in $groups)
					{
						if($group.UserSipAddress -ne $SipAddress)
						{
							$TempArray += @($group)
						}
					}
					$script:groups = $TempArray
				}
			}
		}
	}
	
	#DEBUGGING
	<#
	Write-Host "GROUP ARRAY:"
	$LoopNo = 1
	foreach($group in $groups)
	{
		Write-Host "$LoopNo : " $group.UserSipAddress $group.GroupNumber $group.Version
		$LoopNo++
	}
	#>
}

# Load Groups Array ============================================================
# I WOULD LIKE TO USE THIS METHOD BUT Get-CsGroupPickupUserOrbit IS TOO SLOW... WHAT A SHAME :(
function LoadGroupsSkype 
{
	Write-Host
	Write-Host "Loading Groups from Skype for Business using PowerShell... (This may take some time, perhaps you should get a coffee)" -foreground "green"
	Write-Host
	$script:groups = @()
	
	#THIS METHOD SEEMS TO ALWAYS BE SLOWER...
	<#
	# Get Start Time
	$startDTM = (Get-Date)
	
	$UserGroups = Invoke-Expression "Get-CsUser | Get-CsGroupPickupUserOrbit"
	
	foreach($UserGroup in $UserGroups)
	{
		$GroupUser = $UserGroup.User
		$CPGNumber = $UserGroup.Orbit
		
		$User = Invoke-Expression "Get-CsUser -identity `"$GroupUser`""
		[string]$SipAddress = $User.SipAddress
		[string]$RegistrarPool = $User.RegistrarPool
		
		if($CPGNumber -ne "")
		{
			#Write-Host "FOUND USER IN GROUP: $SipAddress $CPGNumber"
			$script:groups += @(@{"UserPool" = "$UserPoolIdentity";"GroupNumber" = "$CPGNumber";"UserSipAddress" = "$SipAddress"})
		}
	}
	
	# Get End Time
	$endDTM = (Get-Date)
	# Echo Time elapsed
	Write-Host "Group Loop Method Elapsed Time: $(($endDTM-$startDTM).totalseconds) seconds"
	#>
	
	#THIS METHOD IS A LITTLE BIT FASTER...
	
	#Both Methods are super slow...
	# Get Start Time
	$startDTM = (Get-Date)
	
	$Users = Invoke-Expression "Get-CsUser -Filter {EnterpriseVoiceEnabled -eq `$true}"
	
	foreach($User in $Users)
	{
		[string]$SipAddress = $User.SipAddress
		[string]$UserPoolIdentity = $User.RegistrarPool
		$UserGroup = Invoke-Expression "Get-CsGroupPickupUserOrbit -user $SipAddress"
		
		$SipAddress = $SipAddress.Replace("sip:", "")
		[string]$CPGNumber = $UserGroup.Orbit
		
		
		if($CPGNumber -ne "")
		{
			Write-Host "INFO: Found User in Group: $SipAddress $CPGNumber" -foreground "Yellow"
			$script:groups += @(@{"UserPool" = "$UserPoolIdentity";"GroupNumber" = "$CPGNumber";"UserSipAddress" = "$SipAddress";"Version" = 0})
		}
	}
	

	# Get End Time
	$endDTM = (Get-Date)
	# Echo Time elapsed
	Write-Host "Discovery Elapsed Time: $(($endDTM-$startDTM).totalseconds) seconds"
	
	
	#NOTE: Common Area Phones are not supported yet using Skype for Business Commands. When they are, the following could be used:
	<#
	$Phones = Invoke-Expression "Get-CsCommonAreaPhone"
	foreach($Phone in $Phones)
	{
		[string]$SipAddress = $Phone.SipAddress
		[string]$RegistrarPool = $Phone.RegistrarPool
		$UserGroup = Invoke-Expression "Get-CsGroupPickupUserOrbit -user $SipAddress"
		
		$SipAddress = $SipAddress.Replace("sip:", "")
		[string]$CPGNumber = $UserGroup.Orbit
		
		if($CPGNumber -ne "")
		{
			Write-Host "FOUND USER IN GROUP: $SipAddress $CPGNumber"
			$script:groups += @(@{"UserPool" = "$UserPoolIdentity";"GroupNumber" = "$CPGNumber";"UserSipAddress" = "$SipAddress"})
		}
	}
	#>
}

# Update the Group Listbox ============================================================
function UpdateGroupsList {

	#Reset all items Back in Black I hit the sack, I been too long I'm glad to be back...
	foreach($listViewItem in $objGroupsListbox.Items)
	{
		$listViewItem.BackColor = [System.Drawing.Color]::White
		$listViewItem.ForeColor = "Black"
	}
	#Update groups with users in them to Green	
	foreach($Group in $Groups)
	{
		[string]$searchString = $Group.GroupNumber

		[string]$SipAddress = $Group.UserSipAddress

		#Write-Host "USER IN GROUP: $searchString $SipAddress" #DEBUGGING
		
		#$item = $objGroupsListbox.FindItemWithText($searchString)
		#Replaced for accuracy reasons
		$item = $null
		foreach($groupName in $objGroupsListbox.items)
		{
			$groupNameText = $groupName.Text
			#Write-Host "GROUP NAME MATCH $groupName $searchString" #DEBUGGING
			if($groupNameText -eq $searchString)
			{
				$item = $groupName
			}
		}
		
		if ($item -ne $null)
		{
			#Write-Host "FOUND GROUP: " $item $Group.GroupNumber #DEBUGGING
			$item.ForeColor = "Green"
			$item.BackColor = [System.Drawing.Color]::LemonChiffon
		}
	}
	$objGroupsListbox.Refresh()
}

# Update the User Listbox ============================================================
function UpdateUsersList
{
	$objUsersListbox.items.Clear()
	$objGroupsListboxItemSelected = $objGroupsListbox.SelectedItems[0].Text
	$objUsersLabel.Text = "Users in Group $objGroupsListboxItemSelected"
	
	foreach($group in $groups)
	{
		[string] $GroupNumber = $group.GroupNumber
		
		if($GroupNumber -eq $objGroupsListboxItemSelected)
		{
				[void] $objUsersListbox.Items.Add($Group.UserSipAddress)
		}
	}
}

function FilterLyncUsersList
{
	$objLyncUsersListbox.Items.Clear()
	[string] $theFilter = $FilterTextBox.text
	# Add Lync Users ============================================================
	Get-CsUser -Filter {EnterpriseVoiceEnabled -eq $true} | select-object SipAddress | ForEach-Object {if($_.SipAddress -imatch $theFilter){[void] $objLyncUsersListbox.Items.Add($_.SipAddress.ToLower().Replace('sip:',''))}}
	
	if(!$Script:SkypeForBusinessAvailable)
	{
		Get-CsCommonAreaPhone | select-object SipAddress, DisplayName | ForEach-Object {if($_.SipAddress -imatch $theFilter){[void] $objLyncUsersListbox.Items.Add($_.SipAddress.ToLower().Replace('sip:','')); [string]$CommonSipAddress = $_.SipAddress; [string]$CommonDisplayName = $_.DisplayName; write-host "Found Common Area Device: $CommonSipAddress ($CommonDisplayName)" -foreground yellow}}
	}
}

function FindUserGroup([string] $User)
{
	$UserFound = $false
	foreach($group in $groups)
	{
		if($group.UserSipAddress -eq $User)
		{
			$item = $objGroupsListbox.FindItemWithText($group.GroupNumber)
			$item.Selected = $true
			$item.Focused = $true
			UpdateUsersList
			$UserFound = $true
			Write-Host "INFO: User found in group" $group.GroupNumber -foreground "Green"
			break
		}
	}
	if(!$UserFound)
	{
		Write-Host "INFO: User not found in any group." -foreground "Red"
	}
	
}

function RefreshInterface
{
	Write-Host "-------------------------------------------------------------"	
	Write-Host "REFRESHING INTERFACE" -foreground "green"
	Write-Host
	
	$objLyncUsersListbox.Items.Clear()
	
	#Update Orbits
	UpdateSelectedOrbit
	UpdateOrbitsList

	#Select the first Orbit
	$itemsInOrbitsList = $objOrbitsListbox.Items.Count
	if($itemsInOrbitsList -gt 0)
	{$objOrbitsListbox.SelectedIndex = 0}

	#Load Groups
	if($Script:SkypeForBusinessAvailable)
	{
		# Add Lync Users ============================================================
		Get-CsUser -Filter {EnterpriseVoiceEnabled -eq $true} | select-object SipAddress | ForEach-Object {[void] $objLyncUsersListbox.Items.Add($_.SipAddress.ToLower().Replace('sip:',''))}
		LoadGroups
		#LoadGroupsSkype #This method is too slow...
	}
	else
	{
		# Add Lync Users ============================================================
		Get-CsUser -Filter {EnterpriseVoiceEnabled -eq $true} | select-object SipAddress | ForEach-Object {[void] $objLyncUsersListbox.Items.Add($_.SipAddress.ToLower().Replace('sip:',''))}

		Get-CsCommonAreaPhone | select-object SipAddress, DisplayName | ForEach-Object {[void] $objLyncUsersListbox.Items.Add($_.SipAddress.ToLower().Replace('sip:','')); [string]$CommonSipAddress = $_.SipAddress; [string]$CommonDisplayName = $_.DisplayName; write-host "Found Common Area Device: $CommonSipAddress ($CommonDisplayName)" -foreground yellow}
		LoadGroups
	}
	UpdateGroupsList

	#Select the first Group
	$itemsInGroupList = $objGroupsListbox.Items.Count
	if($itemsInGroupList -gt 0)
	{
		$objGroupsListbox.Items[0].Selected = $true
	}

	UpdateUsersList	

	$itemsInLyncUserList = $objLyncUsersListbox.Items.Count
	if($itemsInLyncUserList -gt 0)
	{$objLyncUsersListbox.SelectedIndex = 0}

	$itemsInUserList = $objUsersListbox.Items.Count
	if($itemsInUserList -gt 0)
	{$objUsersListbox.SelectedIndex = 0}
	Write-Host "-------------------------------------------------------------"
}

# Setup Interface  ==============================================================


UpdateOrbitsList

#Select the first Orbit
$itemsInOrbitsList = $objOrbitsListbox.Items.Count
if($itemsInOrbitsList -gt 0)
{$objOrbitsListbox.SelectedIndex = 0}

UpdateSelectedOrbit


#Load Groups
if($Script:SkypeForBusinessAvailable)
{
	# Add Lync Users ============================================================
	Get-CsUser -Filter {EnterpriseVoiceEnabled -eq $true} | select-object SipAddress | ForEach-Object {[void] $objLyncUsersListbox.Items.Add($_.SipAddress.ToLower().Replace('sip:',''))}
	LoadGroups
	#LoadGroupsSkype   #Can't use this method because it's too slow...
}
else
{
	# Add Lync Users ============================================================
	Get-CsUser -Filter {EnterpriseVoiceEnabled -eq $true} | select-object SipAddress | ForEach-Object {[void] $objLyncUsersListbox.Items.Add($_.SipAddress.ToLower().Replace('sip:',''))}

	Get-CsCommonAreaPhone | select-object SipAddress, DisplayName | ForEach-Object {[void] $objLyncUsersListbox.Items.Add($_.SipAddress.ToLower().Replace('sip:','')); [string]$CommonSipAddress = $_.SipAddress; [string]$CommonDisplayName = $_.DisplayName; write-host "Found Common Area Device: $CommonSipAddress ($CommonDisplayName)" -foreground yellow}
	LoadGroups
}
UpdateGroupsList

#Select the first Group
$itemsInGroupList = $objGroupsListbox.Items.Count
if($itemsInGroupList -gt 0)
{
	$objGroupsListbox.Items[0].Selected = $true
}

UpdateUsersList	

$itemsInLyncUserList = $objLyncUsersListbox.Items.Count
if($itemsInLyncUserList -gt 0)
{$objLyncUsersListbox.SelectedIndex = 0}

$itemsInUserList = $objUsersListbox.Items.Count
if($itemsInUserList -gt 0)
{$objUsersListbox.SelectedIndex = 0}



# Activate the form ============================================================
$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()	


# SIG # Begin signature block
# MIIcWAYJKoZIhvcNAQcCoIIcSTCCHEUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUYpRmjQ4yzT80fcXP9ig5yHKX
# ooeggheHMIIFEDCCA/igAwIBAgIQBsCriv7g+QV/64ncHMA83zANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE4MDExMzAwMDAwMFoXDTE5MDQx
# ODEyMDAwMFowTTELMAkGA1UEBhMCQVUxEDAOBgNVBAcTB01pdGNoYW0xFTATBgNV
# BAoTDEphbWVzIEN1c3NlbjEVMBMGA1UEAxMMSmFtZXMgQ3Vzc2VuMIIBIjANBgkq
# hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAukBaV5eP8/bHNonSdpgvTK/2iYj9XRl4
# VzpuJE1fK2sk0ZnjIidsaYXhFpL1LUbUlalxnO7cbWY5ok5bHg0vPx9p8IHHBH28
# xrisz7wcXTTjXMrOL+yynDJYUCMpKV5rMkBn5kJJlLUrY5kcT6Y0fa4HKmvLYAVC
# 6T83mvUvwVs0TlLqY5Dcm/eoVzSmv9Frn3A5WNKxElhhUL2W6LEHdikzCRltk0+e
# g6OXRSYHwulwL+HzcQ+83YEVp/YG9GM+v3Ra4UeuSWaOkt4FQI5JGMlKvhQ3wSu4
# 455xAyj56MTul2FQ1s+j2KI/bvJOMwzO86RDwUC+yZuhh8+IYVObpQIDAQABo4IB
# xTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0OBBYE
# FHXxYdsGH8A4rhw89n7VPGve7xb5MA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAK
# BggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdpY2Vy
# dC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2NybDQu
# ZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUwQzA3
# BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQu
# Y29tL0NQUzAIBgZngQwBBAEwgYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcwAYYY
# aHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8vY2Fj
# ZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNpZ25p
# bmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAQEATz4Xu/3x
# ae3iTkPfYm7uEWpB16eV1Ig+8FMDg6CJ+465oidj2amAjD1n+MwekysJOmcWAiEg
# R7TcQKUpgy5QTTJGSsPm2rwwcBL0jye6hXgs5eD8szZEhdJOnl1txRsdhMtilV2I
# H7X1nQ6S/eRu4WneUUF3YIDreqFYGLIfAobafEEufP7pMk05zgO6lqBM97ee+roR
# eP12IG7CBokmhzoERIDdTjfNEbDtob3OKPKfao2K8MJ079CSoG+NnpieO4CSRQtu
# kaCfg4rK9iCFIksrHq+qSMMRobnVwZq5tDZrkQOjO+lBdL0XWF4nrBavCjs4DjBh
# JHz6nkyqXDNAuTCCBTAwggQYoAMCAQICEAQJGBtf1btmdVNDtW+VUAgwDQYJKoZI
# hvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZ
# MBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNz
# dXJlZCBJRCBSb290IENBMB4XDTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEyMDAwMFow
# cjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVk
# IElEIENvZGUgU2lnbmluZyBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAPjTsxx/DhGvZ3cH0wsxSRnP0PtFmbE620T1f+Wondsy13Hqdp0FLreP+pJD
# wKX5idQ3Gde2qvCchqXYJawOeSg6funRZ9PG+yknx9N7I5TkkSOWkHeC+aGEI2YS
# VDNQdLEoJrskacLCUvIUZ4qJRdQtoaPpiCwgla4cSocI3wz14k1gGL6qxLKucDFm
# M3E+rHCiq85/6XzLkqHlOzEcz+ryCuRXu0q16XTmK/5sy350OTYNkO/ktU6kqepq
# CquE86xnTrXE94zRICUj6whkPlKWwfIPEvTFjg/BougsUfdzvL2FsWKDc0GCB+Q4
# i2pzINAPZHM8np+mM6n9Gd8lk9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB/wQIMAYB
# Af8CAQAwDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHkGCCsG
# AQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29t
# MEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8v
# Y3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqg
# OKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURS
# b290Q0EuY3JsME8GA1UdIARIMEYwOAYKYIZIAYb9bAACBDAqMCgGCCsGAQUFBwIB
# FhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAoGCGCGSAGG/WwDMB0GA1Ud
# DgQWBBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv9JLLgjEt
# UYunpyGd823IDzANBgkqhkiG9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgsfCUpdqgd
# XRwtOhrE7zBh134LYP3DPQ/Er4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJeJIFOEKTu
# P3GOYw4TS63XX0R58zYUBor3nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbHJyqhKSga
# OnEoAjwukaPAJRHinBRHoXpoaK+bp1wgXNlxsQyPu6j4xRJon89Ay0BEpRPw5mQM
# JQhCMrI2iiQC/i9yfhzXSUWW6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDOmTFjPQga
# GLOBm0/GkxAG/AeB+ova+YJJ92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCC
# BmowggVSoAMCAQICEAMBmgI6/1ixa9bV6uYX8GYwDQYJKoZIhvcNAQEFBQAwYjEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBDQS0x
# MB4XDTE0MTAyMjAwMDAwMFoXDTI0MTAyMjAwMDAwMFowRzELMAkGA1UEBhMCVVMx
# ETAPBgNVBAoTCERpZ2lDZXJ0MSUwIwYDVQQDExxEaWdpQ2VydCBUaW1lc3RhbXAg
# UmVzcG9uZGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAo2Rd/Hyz
# 4II14OD2xirmSXU7zG7gU6mfH2RZ5nxrf2uMnVX4kuOe1VpjWwJJUNmDzm9m7t3L
# helfpfnUh3SIRDsZyeX1kZ/GFDmsJOqoSyyRicxeKPRktlC39RKzc5YKZ6O+YZ+u
# 8/0SeHUOplsU/UUjjoZEVX0YhgWMVYd5SEb3yg6Np95OX+Koti1ZAmGIYXIYaLm4
# fO7m5zQvMXeBMB+7NgGN7yfj95rwTDFkjePr+hmHqH7P7IwMNlt6wXq4eMfJBi5G
# EMiN6ARg27xzdPpO2P6qQPGyznBGg+naQKFZOtkVCVeZVjCT88lhzNAIzGvsYkKR
# rALA76TwiRGPdwIDAQABo4IDNTCCAzEwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB
# /wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwggG/BgNVHSAEggG2MIIBsjCC
# AaEGCWCGSAGG/WwHATCCAZIwKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2lj
# ZXJ0LmNvbS9DUFMwggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAAdQBzAGUA
# IABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAYwBvAG4A
# cwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8AZgAgAHQA
# aABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4AZAAgAHQA
# aABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUAZQBtAGUA
# bgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwAaQB0AHkA
# IABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQAIABoAGUA
# cgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZIAYb9bAMV
# MB8GA1UdIwQYMBaAFBUAEisTmLKZB+0e36K+Vw0rZwLNMB0GA1UdDgQWBBRhWk0k
# tkkynUoqeRqDS/QeicHKfTB9BgNVHR8EdjB0MDigNqA0hjJodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDA4oDagNIYyaHR0
# cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcmww
# dwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2Vy
# dC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRENBLTEuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCdJX4b
# M02yJoFcm4bOIyAPgIfliP//sdRqLDHtOhcZcRfNqRu8WhY5AJ3jbITkWkD73gYB
# jDf6m7GdJH7+IKRXrVu3mrBgJuppVyFdNC8fcbCDlBkFazWQEKB7l8f2P+fiEUGm
# vWLZ8Cc9OB0obzpSCfDscGLTYkuw4HOmksDTjjHYL+NtFxMG7uQDthSr849Dp3Gd
# Id0UyhVdkkHa+Q+B0Zl0DSbEDn8btfWg8cZ3BigV6diT5VUW8LsKqxzbXEgnZsij
# iwoc5ZXarsQuWaBh3drzbaJh6YoLbewSGL33VVRAA5Ira8JRwgpIr7DUbuD0FAo6
# G+OPPcqvao173NhEMIIGzTCCBbWgAwIBAgIQBv35A5YDreoACus/J7u6GzANBgkq
# hkiG9w0BAQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBB
# c3N1cmVkIElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMjExMTEwMDAwMDAw
# WjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElE
# IENBLTEwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDogi2Z+crCQpWl
# gHNAcNKeVlRcqcTSQQaPyTP8TUWRXIGf7Syc+BZZ3561JBXCmLm0d0ncicQK2q/L
# XmvtrbBxMevPOkAMRk2T7It6NggDqww0/hhJgv7HxzFIgHweog+SDlDJxofrNj/Y
# MMP/pvf7os1vcyP+rFYFkPAyIRaJxnCI+QWXfaPHQ90C6Ds97bFBo+0/vtuVSMTu
# HrPyvAwrmdDGXRJCgeGDboJzPyZLFJCuWWYKxI2+0s4Grq2Eb0iEm09AufFM8q+Y
# +/bOQF1c9qjxL6/siSLyaxhlscFzrdfx2M8eCnRcQrhofrfVdwonVnwPYqQ/MhRg
# lf0HBKIJAgMBAAGjggN6MIIDdjAOBgNVHQ8BAf8EBAMCAYYwOwYDVR0lBDQwMgYI
# KwYBBQUHAwEGCCsGAQUFBwMCBggrBgEFBQcDAwYIKwYBBQUHAwQGCCsGAQUFBwMI
# MIIB0gYDVR0gBIIByTCCAcUwggG0BgpghkgBhv1sAAEEMIIBpDA6BggrBgEFBQcC
# ARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0
# bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQA
# aABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUA
# dABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkA
# ZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUA
# bAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgA
# aQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAA
# YQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAA
# YgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwEgYDVR0TAQH/
# BAgwBgEB/wIBADB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9v
# Y3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHow
# eDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJl
# ZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDAdBgNVHQ4EFgQUFQASKxOYspkH7R7f
# or5XDStnAs0wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZI
# hvcNAQEFBQADggEBAEZQPsm3KCSnOB22WymvUs9S6TFHq1Zce9UNC0Gz7+x1H3Q4
# 8rJcYaKclcNQ5IK5I9G6OoZyrTh4rHVdFxc0ckeFlFbR67s2hHfMJKXzBBlVqefj
# 56tizfuLLZDCwNK1lL1eT7EF0g49GqkUW6aGMWKoqDPkmzmnxPXOHXh2lCVz5Cqr
# z5x2S+1fwksW5EtwTACJHvzFebxMElf+X+EevAJdqP77BzhPDcZdkbkPZ0XN1oPt
# 55INjbFpjE/7WeAjD9KqrgB87pxCDs+R1ye3Fu4Pw718CqDuLAhVhSK46xgaTfwq
# Ia1JMYNHlXdx3LEbS0scEJx3FMGdTy9alQgpECYxggQ7MIIENwIBATCBhjByMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQg
# Q29kZSBTaWduaW5nIENBAhAGwKuK/uD5BX/ridwcwDzfMAkGBSsOAwIaBQCgeDAY
# BgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBSAOh0s98UHYI0Ui2OThYdpArUdyTANBgkqhkiG9w0BAQEFAASCAQBVIFVtoVWO
# Ilz6a6IifXB6E2s9tCPasYxwh5Q+4XXigDNt4G1qc85AXFbjeOvPvTJKNJRGIxfn
# X6cxBwG2iGjiY4w3xyk/ER+pFII3LdiXTNu29lqbG6fJj8qTwEI7W/9VWsotSfRe
# nzPcHXnuos56qOfQ01RmKNLCdCWwwSsGiGfQWlCJjEHFtXGEfFiWZ51+/o/ySQJf
# cAqIvpbStrZ1dapY56MXdiiIRizK23WUM0/0gSQb6n5NOSzHlRaTboX4sRA9tr6w
# k7UsG6OYIwGpq4cPnIQI5TMgRZIT6x2nE5xAjH9dOBEaGBFQaOIhME7n+CmfEnQg
# j6xdnC2LUSzroYICDzCCAgsGCSqGSIb3DQEJBjGCAfwwggH4AgEBMHYwYjELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBDQS0xAhAD
# AZoCOv9YsWvW1ermF/BmMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTAyMDIwMTQ1MzJaMCMGCSqGSIb3DQEJ
# BDEWBBSib/EKGn47Pf+xEYVwmJmCf8eU1DANBgkqhkiG9w0BAQEFAASCAQBcu2c+
# D5pw9Xk25KLVR9m6vV/cGkRDr9YB2CDMw6N8IsyaNJNpqwv3N/w/gI38lRhG2fQu
# QnPVGJI/r0bcSVX1GvAUwDsq/kB+zCNwo25zlioFggX4aumTwIaRalrBeIzQwuK/
# qzpRupAO346AEgxvGI9QkFoHbwCtGHNZKxNA7CIiSljtgUrNJXL3k9vf66NFl4oN
# 5FfdL043oZ+AGwNfTIpZD2cpuUDIen1bq0sb1GrzH1HDH/Cu+gQm+YteydDvJl2S
# afxQBK9po9qzTzU8uWpbrh9Y92pqG0P+WB/N5SBA+kA7dmXDbKuqw+9bm9UAqt2h
# DB9JOs97MeNZ7Fht
# SIG # End signature block
