namespace SPReports

open System
open System.Collections.Generic
open System.Text.RegularExpressions
open SPReports.DumpMetadata

type webUrl = string
type fileExtension = string
type documentCount = int

type WebsFileExtensionsCountVisitor() = 
    inherit DumpVisitor()
    let mutable webUrl = ""

    // Sparse matrix of webs vs file extensions and counts
    member val Matrix = Dictionary<(webUrl * fileExtension), documentCount>() with get, set
    override __.Visit(w: SPWeb) = webUrl <- w.Url

    override this.Visit(i: SPItem) =
        match i with
        | Document d ->
            // See https://msdn.microsoft.com/en-us/library/office/aa543822.aspx for
            // content type IDs and description of inheritance.
            if d.ContentTypeId.StartsWith("0x0101") then             
                let re = Regex(".*\.(?<fileExtension>[a-zA-Z0-9]+)$")
                let m = re.Match(d.Path)
                let fe = m.Groups.["fileExtension"].Value.ToLower()
                if fe.Length > 0 && d.Version.Major >= 1 then
                    let (exist, value) = this.Matrix.TryGetValue((webUrl, fe))
                    if exist 
                    then this.Matrix.[(webUrl, fe)] <- value + 1
                    else this.Matrix.[(webUrl, fe)] <- 1
        | _ -> ()

module WebsFileExtensionsCount =
    let generateReport dump =
        let v = WebsFileExtensionsCountVisitor()
        DepthFirstTraverser(v).Visit(dump)

        let uniqueFileExtensions =
            v.Matrix.Keys
            |> Seq.map snd
            |> Seq.distinct
            |> Seq.sort
            |> Seq.toList      

        let uniqueWebs =
            v.Matrix.Keys 
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
                        let (exist, value) = v.Matrix.TryGetValue((w, fe))
                        sprintf "%s%d;" acc (if exist then value else 0)) url              
                yield line |]