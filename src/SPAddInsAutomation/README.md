# SharePointAddInsAutomation

Use SharePointAddInsAutomation for automating the updating of add-ins
across site collections.

Below we outline how to identify and update SharePoint add-ins
(formerly known as SharePoint apps). Using
[SPReports](https://github.com/ronnieholm/Bugfree.SharePoint.DiagnosticTools/tree/master/src/SPReports),
first we generate a CSV file with add-ins installed across the
tenant. Next the CSV report is fed into SharePointAddInsAutomation to
automate the app updating flow. Unfortunately, no API for updating is
provided by Microsoft, so we fall back browser automation.

As the running example, we'll use [migrating Nintex to a local
datacenter](https://community.nintex.com/docs/DOC-3921) -- one closer
to where the SharePoint Online tenent is hosted. To complete such a
migration, we must update the Nintex Workflow and Nintex Forms add-ins
to their current versions.

Identifying and updating add-in instances using
SharePointAddInsAutomation works with any add-in, however. The basic
SharePoint dialogs for updating an add-in are the same across any
add-in.

## Identifying which add-ins are installed where

Add-ins are installed on a per web basis so to identify which add-ins
are installed where, webs needs to be recursively searched and any
add-in instance reported. The
[SPReports](https://github.com/ronnieholm/Bugfree.SharePoint.DiagnosticTools/tree/master/src/SPReports)
tool does exactly this:

    .\SPReports.exe --dump-metadata
                    --username rh@bugfree.onmicrosoft.com
                    --password secret
                    --tenant-name bugfree
                    --ouptput-path .\dump.bin
    .\SPReports.exe --webs-add-ins
                    --input-path .\dump.bin
                    --output-path .\webs-add-ins.csv

Dumping the metadata might take hours depending on the size of the
tenant. When it's done, the
[webs-add-ins.csv](Webs-add-ins-sample.csv) file contains the
following columns of interest:

- *Id*. Unique to every installed add-in instance.
- *WebUrl*. Url of the form https://bugfree.sharepoint.com/sites/test denoting the web on which
the add-in is installed.
- *Title*. Either "Nintex Workflow for Office 365" or "Nintex Forms
for Office 365". Those values come in handy when filtering add-ins.
- *Version*. Only applicable to Nintex add-ins. For non-Nintex
add-ins, version is exposed as a custom property. Normally, the CSOM API
doesn't provide a way to get at the version otherwise displayed in the UI.

Thus, now we can filter the CSV for Nintex Workflow and Nintex Forms
add-ins and old versions.

## Updating add-ins using browser automation

Given that no API exists for updating add-ins, and that performing a
manual update on each web (about 300 webs for my initial Nintex case),
is click intensive and time consuming, automation is called for. The
approach taken is to automate the manual steps listed below by remote
controlling the browser:

1. Given a WebUrl such as https://bugfree.sharepoint.com/sites/test,
   navigate to the Site Contents application layout page by appending
   /_layouts/15/viewlsts.aspx to WebUrl.

   <img src="Update-add-in-1.png" height="65%">

   The update link actually contains the Id of the add-in instance such
   it's easy to identify, if present.

2. Upon clicking the update link, a dialog pops up:

   <img src="Update-add-in-2.png" height="55%">

   On the dialog, we want to click the GET IT button.

3. It prompts another dialog to appear where we must trust the add-in:

   <img src="Update-add-in-3.png" height="75%">

   Here we want to click the Trust It button which closes both dialogs
   and takes us back to the Site Contents page and the add-in starts
   updating.

To remote control the browser, we use
[Canopy](https://lefthandedgoat.github.io/canopy), an API build on top
of [Selenium](http://www.seleniumhq.org). Selenium in turn abstracts
away the browser details through drivers such that our automation
logic becomes browser neutral.

## Compiling and running

Requires Visual Studio 2017 with F# language support.

To run SharePointAddInsAutomation, first change the variables marked
as such at the top of Program.fs. Next either re-compile and run the
executable or execute the code in F# interactive.

## Supported platforms

SharePoint on-premise and SharePoint Online.

## Contact

Open an issue on this repository or drop me a mail at mail@bugfree.dk.