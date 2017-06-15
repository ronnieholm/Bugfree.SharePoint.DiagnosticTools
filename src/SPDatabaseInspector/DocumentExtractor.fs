namespace SPDatabaseInspector

open System
open System.Data.SqlClient

module DocumentExtractor =
    type CheckedOutFileMetadata =
        { Id: Guid
          DirName: string
          LeafName: string
      
          // Turns out that some files actually have a size of zero eventhough
          // the actual blob is larger than zero bytes in length. In the test
          // dataset, 0.3% of files had zero length. Note also that the length
          // field is nullable in database, but we have yet to come across a
          // content database holding null value.
          Size: int
          CheckoutUser: string
          CheckoutDate: DateTime

          // Defined as nullable in database, but we haven't yet encountered any
          // actual data with a value of null.
          Extension: string
          TimeCreated: DateTime
          TimeLastModified: DateTime }

    let getCheckedOutFilesMetadata (connection: SqlConnection) =
        let sql = "
            select a.Id, a.DirName, a.LeafName, a.Size, ui.tp_Login,
                   a.CheckoutDate, a.Extension, a.TimeCreated, 
                   a.TimeLastModified
            from dbo.AllDocs a with (nolock)
            left outer join UserInfo ui with (nolock) 
            on ui.tp_SiteID = a.SiteId and
               ui.tp_ID = a.CheckoutUserId
            where a.UIVersionString = '0.1' and
                  a.DeleteTransactionId = 0x and -- not in recycle bin
                  a.Level = 255                  -- checked out"

        use command = new SqlCommand(sql, connection)
        use r = command.ExecuteReader()

        let result = ResizeArray<_>()
        while r.Read() do
            let mapped =
                { Id = r.["Id"] :?> Guid
                  DirName = r.["DirName"] :?> string
                  LeafName = r.["LeafName"] :?> string
                  Size = r.["Size"] :?> int32
                  CheckoutUser = r.["tp_Login"] :?> string
                  CheckoutDate = r.["CheckoutDate"] :?> DateTime
                  Extension = r.["Extension"] :?> string
                  TimeCreated = r.["TimeCreated"] :?> DateTime
                  TimeLastModified = r.["TimeLastModified"] :?> DateTime }
            result.Add(mapped)
        result

    type CheckedOutFileContent =
        { Id: Guid
          Content: byte[] }

    let getCheckedOutFileContent (connection: SqlConnection) (id: Guid) =
        // Once we include the Content column in the projection, actual bytes
        // of the document comes over the wire. Depending on available bandwidth
        // document size, and number of documents this can significantly
        // increase query time.
        let sql = "
            select Id, Content
            from dbo.AllDocStreams a with (nolock)
            where a.Id = @id"

        use command = new SqlCommand(sql, connection)
        command.Parameters.AddWithValue("@id", id) |> ignore
        use r = command.ExecuteReader()
    
        // Multiple elements are returned if the item in question has multiple versions.
        let result = ResizeArray<_>()
        while r.Read() do
            let content = r.["Content"] :?> byte[]
            let mapped =
                { Id = r.["Id"] :?> Guid
                  Content = content }
            result.Add(mapped)
        result