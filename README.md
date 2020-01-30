Lync 2013 / Skype for Business Call Pickup Group Manager
========================================================

            

The Call Pickup Manager Tool offers a Powershell based GUI for configuring and managing Call Pickup Groups within Lync 2013 and Skype for Business.




![Image](https://github.com/jamescussen/lync-2013-skype-for-business-call-pickup-group-manager/raw/master/callpickupgroupmanager2.00_sm.png)



**Features of Lync 2013 / Skype for Business Call Pickup Manager:**


  *  Works with both Lync 2013 using SEFAUtil or with Skype for Business CU1+ using Powershell commands.

  *  View all call pickup group configuration (Orbits, Groups and Users) in one simple interface.

  *  Call Pickup Group Manager discovers call pickup configuration information directly from the Lync / Skype for Business database. This avoids having to poll every user individually using SEFAUtil or Powershell to find their settings. This makes configuration
 discovery much faster than SEFAUtil or Powershell can offer in Lync 2013 or Skype for Business.

  *  Easily Add, Edit or Delete Call Pickup Orbits. 
  *  Group-centric configuration of groups. ie. You can look up groups and see which users are in them rather than looking up each user individually to find their group assignment.

  *  Multi-selectable user list boxes for adding or removing multiple users at once.

  *  Use the 'Find Selected User' button to find which group a user is assigned in.

  *  Use the Filter button to reduce the user list to quickly find the user you want to add.


**Prerequisites:**


  *  Lync 2013: Requires SEFAUtil installed on the system and SQL Dynamic Ports opened in Windows Firewall on all Front End servers.

  *  Skype for Business (CU1+): Requires SQL Dynamic Ports opened in Windows Firewall on all Front End servers.


**Note: There are special configuration items required in order for the SQL discovery method to work and in the case of Lync 2013 you need to configure SEFAUtil. For more details on these
**prerequisites** visit: [http://www.myskypelab.com/2013/10/lync-2013-call-pickup-group-manager.html](http://www.myskypelab.com/2013/10/lync-2013-call-pickup-group-manager.html) **


 


**Releases:**

**1.01 Update:**


  *  Pre-Req check will now look under the default reskit location on all available drives (not just C:)

  *  If SEFAUTIL gives no response (due to an unknown error in SEFAUTIL) the tool will display an error to the user.

  *  Added the Import-Module Lync command in case you run the script from regular Powershell or use the Right Click Run using Powershell method to start the script.


**1.02 Update:**


  *  Added the undocumented '/verbose' flag to the SEFAUtil calls to help with debugging SEFAUtil issues. See post: [http://www.mylynclab.com/2014/04/sefautil-and-lync-2013-call-pickup.html](http://www.mylynclab.com/2014/04/sefautil-and-lync-2013-call-pickup.html)



**1.03 Common Area Phone Update:**


  *  This version has been updated to handle Common Area Phones. Some people reported errors being displayed by the tool when they had manually set (with SEFAUTIL) Group Call Pickup against Common Area Phones (ie. against the SIP URI of the Common Area Device,
 eg: sip:fbcb642b-f5bc-477a-a053-373aef4c00f8@domain.com). As of this version Common Area Phones will be included in the user list, and you can add and remove them from Call Pickup Groups.

  *  User listboxes are now slightly wider to deal with the long SIP addresses of Common Area Phones.

  *  When the tool loads it will display in the Powershell window the SIP Address and Display Name of all common area devices so you can match the (GUID looking) SIP address in the tool to the display name of the device.



**1.04 Scalability Update:**


  *  Now supports window resizing. 
  *  Added Filter on Lync users listbox to cater for deployments that have lots of users.

  *  Script is now signed. 

**1.05 Skype for Business Update**


  *  Now checks that Group number matches a range that exists in one of the Call Pickup Orbits before allowing it to be added to the group list.

  *  Up and Down keys in Orbit listbox now update orbit details properly. 
  *  You can now specify an alternate location of SEFAUTIL.exe in the command line. (Example: .\Lync2013CallPickupManager.1.05 'D:\folder\SEFAUTIL.exe')

  *  Now checks the Skype for Business RESKIT location. 
  *  Put a dividing line between the Orbit creation section and the Group configuration section to try and indicate a divide between the two areas.

  *  The Group list box label now displays the group name being listed up so the user understands better what group the user list is associated with.


**2.00 Major Updates**


  *  When using Skype for Business the new Call Pickup Group Powershell commands (Available in Skype for Business CU1+) for user settings are detected and used instead of SEFAUTIL. This means you don't have to worry about any SEFAUTIL configuration anymore in
 Skype for Business CU1 or higher! 
  *  Unfortunately the Skype for Business Powershell commands have proved to be too slow for the discovery of all users Call Pickup Settings (because 'Get-CsUser | Get-CsGroupPickupUserOrbit' has to iterate through all users taking about 2 seconds per user and
 takes ages). So I have retained and improved the direct SQL discovery method from version 1.0 for both Lync 2013 and Skype for Business.

  *  Changed the way the Groups list box works. It now will be automatically filled with all the available groups from the Orbit ranges assigned in the system. If there is a user in a Group then it will be highlighted in Green text with Yellow background to
 help you find users without looking in empty groups. 
  *  When Orbits are added all the available groups will automatically be added to the available Groups list. Note: when you remove ranges that contain groups with users in them the group will no longer be accessible in the interface, however, the Pickup Group
 will still function. So make sure you remove users from groups before you delete the Orbit range.

  *  Added a refresh button so the data can now be updated from the system whenever you want.

  *  Improved the speed of looking through groups by re-architecting some of the code.

  *  After a user is added or removed from a group the tool does not rescan all user's data again. Whilst this method always ensures the data being displayed is up to date, it's also much slower. This version is optimised for speed :)

  *  Added pretty looking UP and DOWN arrow icons on add and remove buttons to try and clarify operation.

  *  Added the 'Find Selected User' button to find which group a user is assigned to.

  *  Added help tool tips on buttons. 

**2.01 Enhancements**


  *  Added an export to CSV data to the app. 
  *  Fixed an issue with highlighting groups with users in them. 

**2.02 Update**


  *  Now supports Skype for Business 2019 

**For all information on of operation of this tool please visit this link:** 


[http://www.myskypelab.com/2013/10/lync-2013-call-pickup-group-manager.html](http://www.myskypelab.com/2013/10/lync-2013-call-pickup-group-manager.html)


        
    
