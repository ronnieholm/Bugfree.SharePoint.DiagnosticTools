# SPDatabaseInspector

Current capabilities include the extraction of files and metadata
about those files from SharePoint document libraries using direct
database access. Bypassing application layer security has the benefit
of fast and unrestricted access to all of SharePoint. The downside is
that the database schema is mostly undocumented.

## Extract checked out files not previously checked in

Checked out files which haven't previously been checked in exist
because a user added files to a document library but forgot the step
of checking the files in -- perhaps after working on draft versions
for some time. Now only the user can see the files in the document
library. Other users with proper access can navigate to Document
Library Settings and Manage checked out files to browse checked out
files and take ownership.

Using the Manage checked out files function works well through the
browser, but no web service or server-side API exist for retrieving
actual file content. During migration of SharePoint 2007, using
Sharegate or a custom solution, these files will not be migrated. It
makes sense as SharePoint considers those files drafts. But even
drafts may hold valuable information and requires proper handling
during migration.

To invoke extraction from the command-line:

```
USAGE: SPDatabaseInspector.exe [--help] --connection-string <connection> --storage-base-path <path> [--with-table-of-content]

OPTIONS:

    --connection-string <connection>
                          connection string to the SharePoint content database.
    --storage-base-path <path>
                          directory under which extracted files and are stored.
    --with-table-of-content
                          adds metadata file at the root of --storage-base-path containing SharePoint 
						  extracted metadata for each file.
    --help                display this list of options.
```

Example:

```
.\SPDatabaseInspector.exe --connection-string "Server=tcp:<sql-server-host>,1433;Database=<database-name>;Integrated Security=SSPI;Encrypt=True;TrustServerCertificate=True;Connection Timeout=30;" --storage-base-path "c:\temp\output" --with-table-of-content
```

In case the web application is backed by more than one content
database, the command must be run against each content
database. Passing arguments as above, a directory structure resembling
that of SharePoint is created below c:\temp\output. A server-relative
URL to a document library, such as sites/hr/contracts, and a file
inside the contracts document library under 2016/finance/johndoe.docx"
gets created under
c:\temp\output\sites\hr\contracts\2016\finance\johndoe.docx".

Passing the ```--with-table-of-content``` option, a CSV file is
created under c:\temp\output holding metadata of the documents
extracted. The name of the file becomes <database-name>.csv:

- DirName -- sites/hr/contracts/2016/finance
- LeafName -- johndoe.docx
- Size -- 12345
- CheckoutUser -- <domain>\<username>
- CheckoutDate -- 2011-05-10 10:10:51.000
- Extension -- docx
- TimeCreated -- 2011-05-10 10:10:51.000
- TimeLastModified -- 2011-05-10 10:10:51.000
- SHA1 -- hash

Make sure ```--storage-base-path``` is close to the root of the
drive. By limitation of Windows, the path cannot exceed 260 characters
or an exception is thrown.

## Technical details

In any event, basic knowledge about [how the SharePoint content
database is
organized](https://msdn.microsoft.com/en-us/library/hh656481.aspx)
comes in handy. Across versions, the database schema has changed, so
for different versions of SharePoint queries may need to be
adjusted. In general, SharePoint in build around the notion that no
new tables are created after the product is installed. The creation of
site collections, webs, lists, document libraries results only in rows
added to existing tables. Actually, the content database contains a
surprisingly small number of tables given how complex SharePoint
appears through the user interface or API.

A quick warning before running any queries against a live content
database. Always run queries with the ´´´with (nolock)´´´ clause added
to prevent the query from blocking regular SharePoint
operations. SharePoint's application logic may not re-run failed
queries because either the query times exceeded a maximum or because
out query caused a deadlock.

In addition, SharePoint makes use of triggers, which we don't want to
interfere with as that could result in our queries introducing a
content database inconsistency.

Lastly, running our custom queries will interfere with cached query
execution plans and statistics within the database which (for a short
while) can make regular operations slightly slower. The latter is a
minor concern that self-corrects while the first two doesn't. 

The query to identify draft documents is inspired by the stored
procedure called by SharePoint when navigating to Document Library
Settings and Manage checked out files. The procedure is named
[proc_GetListCheckedOutFiles](https://msdn.microsoft.com/en-us/library/hh632446.aspx]
and is called by Microsoft through the CheckedOutFiles property of
SPDocumentLibrary (inside Microsoft.SharePoint.dll).

## Supported platforms

SharePoint 2007.