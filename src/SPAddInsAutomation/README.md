# SharePointAddInsAutomation

SharePointAddInsAutomation automates updating add-ins across site
collections.

The update process consists of two steps. First we identify installed
SharePoint add-ins using the
[SPReports](https://github.com/ronnieholm/Bugfree.SharePoint.DiagnosticTools/tree/master/src/SPReports)
tool. It outputs a CSV file of add-ins installed across the
tenant. Next the CSV file is fed to SharePointAddInsAutomation which
does the updating. As Microsoft provides no API for updating add-ins
we fall back to browser automation.

As an example, we use the updating of Nintex add-ins. Nintex is a good
example a non-singleton type of add-in and [migrating Nintex to a
local datacenter](https://community.nintex.com/docs/DOC-3921) is
preconditioned on add-ins across the tenant being up to date.

Identifying add-ins with SPReports and updating add-ins with
SharePointAddInsAutomation works with any add-in,
however. SharePoint's dialogs for updating add-ins are independent of
the specific add-in.

## Identifying which add-ins are installed where

Add-ins are installed on a per web basis so to identify which add-ins
are installed where, webs must be recursively searched. The
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
following columns of interest to SharePointAddInsAutomation:

- *Id*. Unique to every installed add-in instance.
- *WebUrl*. Url of the form https://bugfree.sharepoint.com/sites/test denoting the web on which the add-in is installed.
- *Title*. For Nintex, the value either reads "Nintex Workflow for Office 365" or "Nintex Forms for Office 365". These values come in handy when filtering add-ins.
- *Version*. Only applicable to Nintex add-ins which expose the version as a custom property. For other add-ins, the CSOM API provides no way to get at the version displayed in the UI.

Thus, now we can filter the Webs-add-ins.csv for Nintex Workflow and
Nintex Forms add-ins not in current versions.

## Updating add-ins using browser automation

Given that no API exists for updating add-ins, and that performing a
manual update on each web is both click intensive and time consuming,
automation is called for. The approach taken by
SharePointAddInsAutomation is to automate the manual steps listed
below by remote controlling the browser:

1. Given a WebUrl such as https://bugfree.sharepoint.com/sites/test,
   navigate to the Site Contents application layout page by appending
   /_layouts/15/viewlsts.aspx to WebUrl.

   <img src="Update-add-in-1.png" height="65%">

   The update link contains the Id of the add-in instance, making the
   link easy to identify among possibly multiple update links on the
   page.

2. Upon clicking the update link, a dialog is displayed:

   <img src="Update-add-in-2.png" height="55%">

   On the dialog, click the GET IT button.

3. Upon clicking the *GET IT* button, another dialog appears asking us
   to trust the add-in:

   <img src="Update-add-in-3.png" height="75%">

   On this dialog, click the *Trust It* button which closes both
   dialogs and takes us back to the Site Contents page where the
   add-in starts updating.

To remote control the browser, we use
[Canopy](https://lefthandedgoat.github.io/canopy), an API build on top
of [Selenium](http://www.seleniumhq.org). Selenium in turn abstracts
away the browser details through drivers, making the automation
browser neutral.

## Compiling and running

Requires Visual Studio 2017 with F# language support.

To run SharePointAddInsAutomation, change the variables marked as such
at the top of Program.fs. Next either re-compile and run the
executable or run the code in F# interactive.

## Supported platforms

SharePoint on-premise and SharePoint Online.

## Contact

Open an issue on this repository or drop me a line at mail@bugfree.dk.