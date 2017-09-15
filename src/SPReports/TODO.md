# TODO

- Implement visitor pattern rather than having each report implement the traversal itself
- Add command-line parameter to control degree of parallism
- Add PowerShell script to exercise all reports
- Add globing support to dump a subset of site collection
- Add flag to only dump a subset of metadata and specify requirements per report
- Add duplicate files report by comparing hashes of document library items
- Finish dump by showing report of runtime, objects indexed, warnings
- Update documentation to reflect new commands
- Rewrite ListSolutions to use binary index
- Use new CSOM API support for accessing file and item version information ([example](https://dev.office.com/blogs/new-sharepoint-csom-version-released-for-Office-365-september-2017))
