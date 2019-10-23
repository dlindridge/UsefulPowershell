# UsefulPowershell #
An assortment of useful (and useless!) Powershell scripts I've written and am willing to share.
####################

### AssaultOnHoth.ps1 ###
A Powershell GUI form to query for all domains in your forest and present you with fields to complete. This is very much a
work in progress though it will slow down for a while if I get distracted by something else (OOH! SHINEY!!)
This is also my first foray in to the world of PSForms so any feedback is welcome.

* Includes AOH_symbol.ico and AOH_background.jpg


### CleanSlateProtocol.ps1 ###
Looks for when a specific user last logged in then deletes the contents of the specified directory if the designated number
of days has passed. If you provide a valid address to send to, warnings will be sent starting 4 days before the event. All
Domain Controllers are queried to get the most recent login date.


### DeleteOlderThan.ps1 ###
Delete files from a directory older than a specified number of days.
Replaces FORFILES DOS command (depricated).


### FailureToCommunicate.ps1 ###
Lock an AD account by attempting to login with/use it one more than the Domain Lockout Policy allows.


### IpInfo.ps1 ###
Get the public IP address of the machine this is run on and send it as an email or place a text file somewhere.


### LamentConfiguration.ps1 ###
Queries your domain to find locked users and the Domain Controller that registered the most recent logon (this should in 
theory put it closest to the user) then presents you with options to unlock the account.

* Includes LC_symbol.ico and LC_background.jpg


### RenameFiles.ps1 ###
Find files in a root directory (Recursivly) and rename or copy them to the same relative directory.


### Shenanigans.ps1 ###
This script is a bit monolithic, but it is filled with routine maintenance tasks you should be running for Active Directory
anyways. You can pick which actions you want to run by setting the correct option in the Enable Script Actions section. I've
tried to comment this as throughly as practical so it should make sense as you go through it. It was written as a consolidation
of multiple other scripts so there may be some leftover oddities, but it does work as of writing this. Searches that could be
impacted by querying a specific Domain Controller (i.e. LastLogon) compensate by searching ALL Domain Controllers in the
specified domain and taking the most recent result.

To use this script make sure that service account you run this under has permission to change objects in the OUs specified. This
script must be run from a server that has the Active Directory Module for Powershell or RSAT installed. I recommend setting it up
as a daily scheduled task - I run mine daily at 3am.

"Script Options" section should handle all of your customization needs, but pay particular attention to the email notification
in the "Domain Password Notice" section if your preferred method of changing passwords is not on the client PC or your password
requirements are different than the standard Complexity options in Active Directory.


### SmtpTest.ps1 ###
Simple SMTP relay test to validate your relay settings by sending an email to address specified.


### SortFiles.ps1 ###
Sort all files in a root directory to new folders in the same directory based on the original file name.


### ThirtySevenFlairs.ps1 ###
Queries your domain for any DFSR locations then builds a report (display or email output) with the backlog.
