# CHANGELOG

## SPReports 1.5.0 (2017-09-17)

* Updated Nuget packages to latest version
* Implemented visitor pattern rather than having each report do its own traversal
* Removed support for list-solutions-from-tenant as SPO no longer supports code-based solutions

## SPReports 1.4.1 (2017-08-09)

* Adjust for removed SMLastModifiedDate property on WorkflowDefinition ([#1](https://github.com/ronnieholm/Bugfree.SharePoint.DiagnosticTools/issues/1))
* Prevent dumping tenant metadata from failing due to intermittent communication issues ([#2](https://github.com/ronnieholm/Bugfree.SharePoint.DiagnosticTools/issues/2))

## SPReports 1.3.0 (2017-05-30)

* Updated Nuget packages to latest version
* Updated F# projects to F# Core 4.1
* When traversing tenant, include metadata about AddIns in the dump
* When traversing tenant, include metadata about Workflows in the dump
* Switch to Argu for more robost command-line parsing

##  SPReports 1.2.0 (2016-09-12)

* Include item version in dump

##  SPReports 1.1.0 (2016-08-24)

* Added --dump-metadata and --webs-file-extensions-count options

## SPReports 1.0.0 (2016-08-11)

* Initial release