# SPTransferSpeed

Download compiled version of tools under [Releases](https://github.com/ronnieholm/Bugfree.SharePoint.DiagnosticTools/releases).

## When to use it

For diagnosing network-related issues with the transfer speed between a user's
machine and SharePoint. Oftentimes the issue isn't with SharePoint but the 
hardware and software in between. Speed is a subjective matter and having a 
tool provide quantifiable measurements may help confirm an issue.

## How to use it

After compiling the solution, create a document library for storing the test
files. Then run SPTransferSpeed as follows (username and password are for 
SharePoint Online only):

```
# SPTransferSpeed.exe documentItemUrl sizeInMB [username] [password]
%> .\SPTransferSpeed.exe https://bugfree.sharepoint.com/sites/transferspeed/tests/test1 1024 rh@bugfree.onmicrosoft.com password

1024 MB uploaded in 280.9 seconds at 30.6 Mbit/s
1024 MB downloaded in 95.6 seconds at 89.8 Mbit/s
```

This causes SPTransferSpeed to create a file by the name test1 of 1024 MB 
bytes inside the tests document library. Once the upload and download
completes, a report is displayed.

## How it works

The application makes use of the streaming upload and download capabilities 
of CSOM to transfer a file with random content from the client to a SharePoint 
document library. It then downloads the file, keeping track of how long each 
operation takes.

## Supported platforms

SharePoint 2013 on-prem and SharePoint Online.

## See also

[Informally speed testing SharePoint Online](http://bugfree.dk/blog/2015/03/28/informally-speed-testing-sharepoint-online/)
