namespace SPReports

open SPReports.DumpMetadata

module WebsAddIns =
    let lines = ResizeArray<string>()

    let rec visitWeb level (w: SPWeb) =        
        printfn "W: %s%s" (String.replicate level " ") w.Url
        match w.AddInInstances with
        | Some instances ->
            for a in instances do
                lines.Add(sprintf "%s; %s; %s; %s; %s; %s; %s" 
                            (a.Id.ToString()) 
                            w.Url
                            (a.ProductId.ToString()) 
                            a.Title 
                            (match a.Version_ with | Some v -> v | None -> "N/A") 
                            a.AppPrincipalId 
                            a.AppWebFullUrl)
        | None -> ()

        match w.Webs with
        | Some webs -> webs |> Seq.iter (visitWeb (level + 2))
        | None -> ()

    let visitSiteCollection (sc: SPSiteCollection) =     
        printfn "SC: %s" sc.Url
        match sc.RootWeb with
        | Some sc' -> visitWeb 2 sc'
        | None -> ()

    let generateReport dump =
        dump |> Seq.iter visitSiteCollection
        [| yield "Id; WebUrl; ProductId; Title; Version (Nintex); AppPrincipalId; AppWebFullUrl"
           for s in lines do yield s |]