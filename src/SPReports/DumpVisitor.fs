namespace SPReports

open SPReports.DumpMetadata

type DumpVisitor() =
    abstract Visit: SPSiteCollection -> unit
    abstract Visit: SPWeb -> unit
    abstract Visit: SPList -> unit
    abstract Visit: SPItem -> unit
    abstract Visit: unit -> unit
    abstract Visit: AddIn -> unit
    abstract Visit: Workflow -> unit

    override __.Visit(sc: SPSiteCollection) = ()
    override __.Visit(w: SPWeb) = ()
    override __.Visit(l: SPList) = ()
    override __.Visit(i: SPItem) = ()
    override __.Visit() = ()
    override __.Visit(i: AddIn) = ()
    override __.Visit(wf: Workflow) = ()

type DepthFirstTraverser(visitor: DumpVisitor) =
    let visitListItem (li: SPItem) =
        visitor.Visit(li)

    let visitList level (l: SPList) =
        printfn "L: %s%s" (String.replicate level " ") l.Title
        visitor.Visit(l)

        match l.Items with
        | Some items -> items |> Array.iter visitListItem
        | None -> ()

    let rec visitWeb level (w: SPWeb) =        
        printfn "W: %s%s" (String.replicate level " ") w.Url
        visitor.Visit w

        match w.AddInInstances with
        | Some instances -> instances |> Array.iter visitor.Visit
        | None -> ()

        match w.Workflows with
        | Some workflows -> workflows |> Array.iter visitor.Visit
        | None -> ()

        match w.Lists with
        | Some lists -> lists |> Array.iter (visitList (level + 2))
        | None -> ()

        match w.Webs with
        | Some webs -> webs |> Seq.iter (visitWeb (level + 2))
        | None -> ()

    let visitSiteCollection sc =
        printfn "SC: %s" sc.Url
        visitor.Visit sc
        match sc.RootWeb with
        | Some w -> visitWeb 2 w
        | None -> ()

    member __.Visit(dump: seq<SPSiteCollection>) =
        dump |> Seq.iter visitSiteCollection
        ()