# SPAccessMethodLatency

The purpose of this tool is to compare and contrast the write-read
latency across SharePoint's data access methods. Through periodic
measurements, the results can reveal performance issues and/or a
sub-optimal [BLOB cache
configuration](https://blogs.msdn.microsoft.com/spses/2012/08/28/blob-cache-in-sharepoint).

With each measurement, the tool reports on the following parameters
(column in output):

- *Run*. Monotonically increasing measure identifier.

- *Write*. The write time as observed by the tool and written as
  document library file content for later delta computation. Compared
  to SharePoint's created by or modified by timestamp, the tool
  control this one.

- *CSOM*. Delta of file content write time read back using CSOM and
   original write time.

- *HTTP1 (on-prem only)*. Delta of file content write time read by
   issuing an HTTP request to download the same URL as when displaying
   the file's content in the browser. In case the browser experiences
   latency due to server-side caching (such as BLOB caching), it's
   reflected by a negative HTTP1 value indicating stale content.

- *HTTP2 (on-prem only)*. Same as HTTP1 but with
   ```version=someGuid``` added as a query string parameter. The
   parameter's name and value are made up for the purpose of
   generating a unique URL to prevent server and/or client-side
   caching.

- *REST1*. Delta of file content write time and SharePoint's
   modified. This measure includes network round-trip time, server
   processing time, and database query time, and is an indirect
   measure of how busy the platform is.

- *REST1*. Delta of file content write time read by issuing an OData
   REST openbinarystream request and original write time.

## Example

Here's an example of the tool running on-prem with BLOB cache enabled
for the JS file extension. Every 5 seconds a new measurement is taken
(username and password must only be supplied with SharePoint Online):

    # SPAccessMethodLiveliness.exe destinationLibraryUrl pingInterval [username] [password]
    %> .\SPAccessMethodLatency.exe https://acme.com/sites/liveliness/library/some.js 5

From the HTTP1 columns, we say that the platform is configured with a
BLOB cache for the JS file extension, and the BLOB cache expires
entries after one day (86,400 seconds).

The numbers in the CSOM, HTTP1, HTTP2, REST1, and REST2 columns is the
drift in seconds between actual item and what SharePoint returns based
on the access method.

    Run   Write             CSOM     HTTP1     HTTP2    REST1      REST2
    0     01/12/2016 15:47     0         0         0        0          0
    1     01/12/2016 15:47     0       -30         0        0          0
    2     01/12/2016 15:48     0       -60         0        0          0
    ...
    2878  02/12/2016 15:46     0    -86343         0      0.1          0
    2879  02/12/2016 15:46     0    -86373         0      0.1          0
    2880  02/12/2016 15:47     0         0         0      0.1          0
    2881  02/12/2016 15:47     0       -30         0      0.1          0
    2882  02/12/2016 15:48     0       -60         0      0.1          0
    ...
    5756  03/12/2016 15:46     0    -86343         0      0.8          0
    5757  03/12/2016 15:46     0    -86373         0      0.8          0
    5758  03/12/2016 15:47     0         0         0      0.8          0
    5759  03/12/2016 15:47     0       -30         0      0.9          0
    5760  03/12/2016 15:48     0       -60         0      0.9          0

We can tell from HTTP2 that appending a query parameter to the URL
causes SharePoint to not use BLOB caching. BLOB caching is solely used
for browser based requests and has no effect on results returned
through CSOM or REST.

For REST1, a little less than one second may sound like a long
time. However, SharePoint Online typically shows higher values. Values
in the three to four second range is common. That's probably why
editing items sometimes results in a conflict saying the item was
modified by another user. That user is likely yourself because on save
the user interface loads faster than the value of REST1 causing its
optimistic concurrency check to load stale data.

    %> .\SPAccessMethodLatency.exe https://bugfree.sharepoint.com/sites/liveliness/library/some.js 5 rh@bugfree.onmicrosoft.com password | tee Log.txt

    Run     Write                   CSOM    REST1   REST2
    0       1/13/2017 8:31:18 AM    0       1.9     0
    1       1/13/2017 8:31:23 AM    0       0.9     0
    2       1/13/2017 8:31:28 AM    0       1.9     0
    3       1/13/2017 8:31:33 AM    0       1.9     0
    4       1/13/2017 8:31:38 AM    0       0.9     0
    5       1/13/2017 8:31:43 AM    0       1.9     0

SharePoint Online doesn't support BLOB caching, by the way.

## Supported platforms

SharePoint 2013 and SharePoint Online.