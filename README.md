# UsefulPowershell #
An assortment of useful (and useless!) Powershell scripts I've written and am willing to share.
Check back on the regular because I'll post updates without warning as something tickles my fancy or I figure out a solution to something.

symbol.ico and background.jpg are included for optional decoration on GUI forms. Substitute appropriate images if you want something else.
####################

### CleanSlateProtocol.ps1 ### (Command Line)
Looks for when a specific user last logged in then deletes the contents of the specified directory if the designated number
of days has passed. If you provide a valid address to send to, warnings will be sent starting 4 days before the event. All
Domain Controllers are queried to get the most recent login date.


### DangItBobby.ps1 ### (GUI Form)
This is a GUI version of the original script by bluesole (https://bluesoul.me). I've done a little cleanup, but most of my
effort was put into getting it into a Form. It looks for a user logged into an available computer on your network and
allows you to disable active NICs. Useful if you are getting hit by a Cryptolocker attack.


### DeleteOlderThan.ps1 ### (Command Line)
Delete files from a directory older than a specified number of days.
Replaces FORFILES DOS command (depricated).


### DFSRepReport.ps1 ### (Command Line)
Queries your domain for any DFSR locations then builds a report (display or email output) with the backlog.


### FSMOQuery.ps1 ### (GUI Form)
Get FSMO Role assignments and other stuff for the selected domain.


### FileSearch.ps1 ### (Command Line)
Recursively searches directory to find a file or folder either by specific name or by pattern.


### FindEventID.ps1 ### (Command Line)
Looks through the event log to find a specific Event ID number and outputs results to file choice (CSV or TXT).


### IpInfo.ps1 ### (Command Line)
Get the public IP address of the machine this is run on and send it as an email or place a text file somewhere.


### LockADAccount.ps1 ### (Command Line)
Lock an AD account by attempting to login with/use it one more than the Domain Lockout Policy allows.


### NewUser.ps1 ### (GUI Form)
A Powershell GUI form to query for all domains in your forest and present you with fields to complete. This is very much a
work in progress though it will slow down for a while if I get distracted by something else (OOH! SHINEY!!)
This is also my first foray in to the world of PSForms so any feedback is welcome.


### OutlookSigUpdate.ps1 ### (Command Line)
Using an HTML template for an email signature, customize it for each user in a specified source CSV to use as a
signature block with their local Outlook client. Ideal for running as a login script to keep signature information
current from Active Directory.


### OWASigUpdate.ps1 ### (Command Line)
Using an HTML template for an email signature, customize it for each user in a specified source CSV to use as a
signature block with OWA. Pulls data directly from Active Directory and connects to Office 365/Exchange Online
to apply the signature and set it as default.


### RandomPasswordGenerator.ps1 ### (Command Line)
Gets all enabled accounts in the domain (at the Search OU Root) and generates a random 15 character password that meets
Microsoft AD password requirements. Generates an output CSV with the passwords assigned to each user.


### RenameFiles.ps1 ### (Command Line)
Find files in a root directory (Recursivly) and rename or copy them to the same relative directory.


### Shenanigans.ps1 ### (Command Line)
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


### SmtpTest.ps1 ### (Command Line)
Simple SMTP relay test to validate your relay settings by sending an email to address specified.


### SortFiles.ps1 ### (Command Line)
Sort all files in a root directory to new folders in the same directory based on the original file name.


### UnlockUser.ps1 ### (GUI Form)
Queries your domain to find locked users and the Domain Controller that registered the most recent logon (this should in 
theory put it closest to the user) then presents you with options to unlock the account.
