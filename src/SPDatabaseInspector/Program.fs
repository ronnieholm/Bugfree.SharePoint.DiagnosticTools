open System
open System.Data.SqlClient
open System.IO
open System.Text
open Argu
open SPDatabaseInspector.DocumentExtractor

type CLIArguments =
    | [<Mandatory>] Connection_string of connection: string
    | [<Mandatory>] Storage_base_path of path: string
    | With_table_of_content
with
    interface IArgParserTemplate with
        member s.Usage =
            match s with
            | Connection_string _ -> "connection string to the SharePoint content database."
            | Storage_base_path _ -> "directory under which extracted files and are stored."
            | With_table_of_content _ -> "adds metadata file at the root of --storage-base-path containing SharePoint extracted metadata for each file."

let parsedArgsOrException args =
    try
        let parser = ArgumentParser.Create<CLIArguments>()
        Choice1Of2(parser.Parse(args))
    with
    | ex -> Choice2Of2(ex)

let formatDateTime (dt: DateTime) = dt.ToString("yyyy-MM-dd HH:mm:ss.fff")

[<EntryPoint>]
let main argv =
    let csv = ResizeArray<_>()
    csv.Add("DirName;LeafName;Size;CheckoutUser;CheckoutDate;Extension;TimeCreated;TimeLastModified;SHA1")

    match parsedArgsOrException argv with
    | Choice1Of2(args) ->        
        use connection = new SqlConnection(args.GetResult <@ Connection_string @>)
        connection.Open()

        let checkedOutFiles = getCheckedOutFilesMetadata connection |> Seq.toArray (*|> Array.take 10*)
        let estimatedSize = (checkedOutFiles |> Array.fold (fun acc cur -> acc + int64(cur.Size)) 0L)
        printfn "Found %d checked out files of an estimated size of %d MB" (checkedOutFiles.Length) (estimatedSize / (1024L * 1024L))

        checkedOutFiles |> Array.iteri (fun idx f ->
            let file = getCheckedOutFileContent connection (f.Id) |> Seq.exactlyOne
            printfn "[%d/%d] %s (%.2f MB)" (idx + 1) (checkedOutFiles |> Array.length) f.LeafName (float(f.Size) / (1024. * 1024.))

            let relativeFolderPath = f.DirName.Replace('/', '\\')
            let absoluteFolderPath = Path.Combine(args.GetResult <@ Storage_base_path @>, relativeFolderPath)
            let absoluteFilePath = Path.Combine(absoluteFolderPath, f.LeafName)

            let hash = System.Security.Cryptography.HashAlgorithm.Create("SHA1").ComputeHash(file.Content)
            let prettyHash = BitConverter.ToString(hash).Replace("-", "").ToLower()

            Directory.CreateDirectory(absoluteFolderPath) |> ignore
            File.WriteAllBytes(absoluteFilePath, file.Content)
            csv.Add(sprintf "%s;%s;%d;%s;%s;%s;%s;%s;%s" f.DirName f.LeafName f.Size f.CheckoutUser (formatDateTime(f.CheckoutDate)) f.Extension (formatDateTime(f.TimeCreated)) (formatDateTime(f.TimeLastModified)) prettyHash))
        match (args.TryGetResult <@ With_table_of_content @>) with
        | Some _ -> 
            let builder = SqlConnectionStringBuilder(args.GetResult <@ Connection_string @>)
            let databaseName = builder.InitialCatalog
            File.WriteAllLines(Path.Combine(args.GetResult <@ Storage_base_path @>, sprintf "%s.csv" (builder.InitialCatalog)), csv, Encoding.UTF8)
        | None -> ()
    | Choice2Of2(ex) ->
        printfn "%s" ex.Message

    0
