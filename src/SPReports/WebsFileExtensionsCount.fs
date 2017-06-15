namespace SPReports

open System
open System.Collections.Generic
open System.Text.RegularExpressions
open SPReports.DumpMetadata

module WebsFileExtensionsCount =
    type webUrl = string
    type fileExtension = string
    type documentCount = int

    // Sparse matrix of webs vs file extensions and counts
    let matrix = Dictionary<(webUrl * fileExtension), documentCount>()

    let rec visitWeb level (w: SPWeb) =        
        printfn "W: %s%s" (String.replicate level " ") w.Url          
        
        match w.Lists with
        | Some lists ->
            for l in lists do
                match l.Items with
                | Some items ->
                    items
                    |> Array.iter (fun li -> 
                        match li with
                        | Document d ->
                            // See https://msdn.microsoft.com/en-us/library/office/aa543822.aspx for
                            // content type IDs and description of inheritance.
                            if d.ContentTypeId.StartsWith("0x0101") then             
                                let re = Regex(".*\.(?<fileExtension>[a-zA-Z0-9]+)$")
                                let m = re.Match(d.Path)
                                let fe = m.Groups.["fileExtension"].Value.ToLower()
                                if fe.Length > 0 && d.Version.Major >= 1 then
                                    let (exist, value) = matrix.TryGetValue((w.Url, fe))
                                    if exist 
                                    then matrix.[(w.Url, fe)] <- value + 1
                                    else matrix.[(w.Url, fe)] <- 1
                        | _ -> ())
                    | None -> ()
        | None -> ()

        match w.Webs with
        | Some webs -> webs |> Seq.iter (visitWeb (level + 2))
        | None -> ()

    // todo: Why isn't this called? Include in generateReport?
    let visitSiteCollection (sc: SPSiteCollection) =     
        printfn "SC: %s" sc.Url
        match sc.RootWeb with
        | Some sc' -> visitWeb 2 sc'
        | None -> ()

    let generateReport dump =
        let uniqueFileExtensions =
            matrix.Keys
            |> Seq.map snd
            |> Seq.distinct
            |> Seq.sort
            |> Seq.toList      

        let uniqueWebs =
            matrix.Keys 
            |> Seq.map fst
            |> Seq.distinct
            |> Seq.sort
            |> Seq.toList

        [| yield sprintf "%s" (uniqueFileExtensions |> Seq.fold (sprintf "%s;'%s'") "Url")
           for w in uniqueWebs do
                let url = sprintf "%s;" w
                let line = 
                    uniqueFileExtensions 
                    |> Seq.fold (fun acc fe ->
                        let (exist, value) = matrix.TryGetValue((w, fe))
                        sprintf "%s%d;" acc (if exist then value else 0)) url              
                yield line |]