# UsefulPowershell
An assortment of useful (and useless!) Powershell scripts I've written and am willing to share.


### AssaultOnHoth.ps1 ###
A Powershell GUI form to query for all domains in your forest and present you with fields to complete. 
This is very much a work in progress though it will slow down for a while if I get distracted by something else (OOH! SHINEY!!)
This is also my first foray in to the world of PSForms so any feedback is welcome.

* Includes AOH_symbol.ico and AOH_background.jpg


### IpInfo.ps1 ###
Get the public IP address of the machine this is run on and send it as an email or place a text file somewhere.


### ThirtySevenFlairs.ps1 ###
Queries your domain for any DFSR locations then builds a report (display or email output) with the backlog.


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
