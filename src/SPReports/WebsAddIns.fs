namespace SPReports

open SPReports.DumpMetadata

type WebsAddInsVisitor() = 
    inherit DumpVisitor()
    let mutable webUrl = ""
    member val Lines = ResizeArray<string>() with get, set

    override __.Visit(w: SPWeb) = webUrl <- w.Url
    override this.Visit(i: AddIn) = 
        this.Lines.Add(sprintf "%s; %s; %s; %s; %s; %s; %s" 
                        (i.Id.ToString()) 
                        webUrl
                        (i.ProductId.ToString()) 
                        i.Title 
                        (match i.Version_ with | Some v -> v | None -> "N/A") 
                        i.AppPrincipalId 
                        i.AppWebFullUrl)

module WebsAddIns = 
    let generateReport dump =
        let v = WebsAddInsVisitor()
        DepthFirstTraverser(v).Visit(dump)        
        v.Lines.Insert(0, "Id; WebUrl; ProductId; Title; Version (Nintex); AppPrincipalId; AppWebFullUrl")
        v.Lines.AsReadOnly()