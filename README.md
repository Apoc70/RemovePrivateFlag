## Synopsis

RemovePrivatFlags.exe will remove the privat flags from all messages in a mailbox (appointments it self will not be altered).

## Why should I use this

If a private meeting request is forwarded or answered the message it self is also marked as private. This is fine for personal
mailboxes but you will maybe get problems if you use shared mailboxes or after migrate data from public folders into a mailbox.
This utility is designed for run once until a migration.

## Installation

Simply copy all files to a location where you are allowed to run it and of course the Exchange servers are reachable.

## Requirements
* Exchange Server 2013 (Tested with CU15, maybe it will work with Exchange 2007/2010 as well)
* Application Impersonisation rights
* Microsoft.Exchange.WebServices.dll, log4net.dll (are provided in the repository and also in the binarys)

## Usage

_RemovePrivateFlags.exe -mailbox user@domain.com [-logonly true] [-foldername "Inbox"]_
Search through the mailbox and ask for changing a item if -logonly is not set to true. If -foldername is given the folder path are compared to the folder name.
If -logonly is set to true only a log will be created.

_RemovePrivateFlags.exe -mailbox user@domain.com [-foldername "Inbox"] [-noconfirmation true]_
Search through the mailbox, if -noconfirmation is set to true all items will be altered without confirmation.

# Parameters
* not optional: -mailbox user@domain.com
* optional: -logonly true
* optional: -foldername "Inbox"
* optional: -noconfirmation true


## License

MIT License
