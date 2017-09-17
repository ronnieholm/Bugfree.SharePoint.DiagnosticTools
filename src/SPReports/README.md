# SPReports

Download compiled version of tools under [Releases](https://github.com/ronnieholm/Bugfree.SharePoint.DiagnosticTools/releases).

Generates custom reports based on metadata in SharePoint by first
collecting and dumping relevant metadata to a file and then quering it
offline. SPReports comes with a few pre-defined reports, but using any
.NET language, new reports can be crafted from the metadata dump.

```
USAGE: SPReports.exe [--help] [<subcommand> [<options>]]

SUBCOMMANDS:

    --dump-metadata <options>
                          dump metadata.
    --webs-file-extensions-count <options>
                          analyze dump and summarize files by extension and count.
    --webs-add-ins <options>
                          analyze dump and create webs by add ins table.
    --webs-workflows <options>
                          analyze dump and create webs by workflow instances table.

    Use 'SPReports.exe <subcommand> --help' for additional information.

OPTIONS:

    --help                display this list of options.
```

## --dump-metadata

Dumps metadata from across site collections, webs, lists, and
workflows for later consumption by reporting commands. Instead of
combining collection and dumping, which carries a heavy runtime cost,
with reporting, the latter is kept separate for speed and
reproducability.

Example:

    .\SPReports.exe --dump-metadata --username rh@bugfree.onmicrosoft.com --password secret --tenant-name bugfree --ouptput-path .\dump.bin

Collects metadata using the username and password provided. All site
collections, webs, lists, and workflows within the bugfree tenant are
traversed, and output is written to a binary file named dump.bin.

The user provided must be a tenant administrator.

## --webs-file-extensions-count

Generates a CSV report of webs by file extension count in the form of
a CSV file. This is useful for getting an idea of which types of files
users are storing. Perhaps files with some extensions shouldn't go
into SharePoint but another application, e.g., CAD drawings or GIS
data.

Example:

    .\SPReports.exe --webs-file-extensions-count --input-path .\dump.bin --output-path .\dump.csv

Outputs a CSV file containing a matrix enumerating webs and file
extensions extracted from the dump file. The value of the (web,
extension) element in the matrix is the number of files within the web
with that extension.

## --webs-add-ins

Generates a CSV report containing apps installes on all webs. For
Nintex Workflow and Nintex Forms in particular, the report also
includes their version numbers. For instance, this enables upgrading
the Nintex Workflow and Forms app to a version supported when moving
from one Nintex datacenter to another.

Example:

    .\SPReports.exe --webs-add-ins --input-path .\dump.bin --output-path .\dump.csv

## --webs-workflows

Generates a CSV report containing workflow instances across all webs
with a column descriminating SharePoint's build-in workflow instances
from Nintex instances.

Example:

    .\SPReports.exe --webs-workflows --input-path .\dump.bin --output-path .\dump.csv

## Supported platforms

SharePoint 2013 and SharePoint Online.