## Note

This is the beta branch. Build can failing or build executables can contain errors.

Use at your own risk.

## Synopsis

RemovePrivateFlags.exe will remove the private flags from all messages in an Exchange Server mailbox (appointments will not be altered).

## Why should I use this

If a private meeting request is forwarded or answered the message it self is also marked as private. This is fine for personal
mailboxes but you will maybe get problems if you use shared mailboxes or after migrate data from public folders into a mailbox.

This utility is designed for run once until a migration.

## Installation

Simply copy all files to a location where you are allowed to run it and of course the Exchange servers are reachable.

## Requirements
* Exchange Server 2013/2016 (Tested with CU15, maybe it will work with Exchange 2007/2010 as well)
* Application Impersonation Rights if you want to change items on other mailboxes than yours
* Microsoft.Exchange.WebServices.dll, log4net.dll (are provided in the repository and also in the binaries)

## Usage
```
RemovePrivateFlags.exe -mailbox user@domain.com [-logonly] [-foldername "Inbox"]
```

Search through the mailbox and ask for changing a item if -logonly is not set to true. If -foldername is given the folder path are compared to the folder name.
If -logonly is set to true only a log will be created.


```
RemovePrivateFlags.exe -mailbox user@domain.com [-foldername "Inbox"] [-noconfirmation]
```

Search through the mailbox, if -noconfirmation is set to true all items will be altered without confirmation.

# Parameters
* mandatory: -mailbox user@domain.com

Mailbox which you want to alter.

* optional: -logonly 

Items will only be logged.

* optional: -foldername "Inbox"

Will filter the items to the Folderpath

* optional: -noconfirmation

Messages will be set to normal without confirmation.

* optional: -ignorecertificate

Ignore certificate errors. Interesting if you connect to a lab config with self signed certificate.

* optional: -impersonate

If you want to alter a other mailbox than yours set this parameter.

* optional: -user user@domain.com

If set together with -password this credentials would be used. Elsewhere the credentials from your session will be used.

* optional: -password "Pa$$w0rd"
* optional: -url "https://server/EWS/Exchange.asmx"

If you set an specific URL this URL will be used instead of autodiscover. Should be used with -ignorecertificate if your CN is not in the certficate.

* optional: -allowredirection

If your autodiscover redirects you the default behaviour is to quit the connection. With this parameter you will be connected anyhow (Hint: O365)



## License

MIT License
