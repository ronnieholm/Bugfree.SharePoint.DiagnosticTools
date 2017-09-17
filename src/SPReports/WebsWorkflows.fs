namespace SPReports

// Loosely inspired by
//   https://community.nintex.com/community/build-your-own/nintex-for-office-365/blog/2017/02/20/migrate-your-tenant-to-another-datacenter-1-of-3
//   https://community.nintex.com/community/build-your-own/nintex-for-office-365/blog/2017/02/22/migrate-your-tenant-to-another-datacenter-2-of-3-or-where-is-the-workflow-app
//   https://community.nintex.com/community/build-your-own/nintex-for-office-365/blog/2017/04/20/migrate-your-tenant-to-another-datacenter-3-of-3-where-are-all-the-workflows

open System
open System.Collections.Generic
open SPReports.DumpMetadata

type WebsWorkflowVisitor() = 
    inherit DumpVisitor()
    let mutable webUrl = ""
    let mutable (listIdToString: Guid option -> string) = fun __ -> ""
    member val Lines = ResizeArray<string>() with get, set

    override __.Visit(w: SPWeb) =
        let lookupListById guid =
            match w.Lists with
            | Some lists -> lists |> Array.tryFind (fun l -> l.Id = guid)
            | None -> None

        webUrl <- w.Url      
        listIdToString <-
            fun id ->
                match id with
                | Some id ->
                    lookupListById id
                    |> Option.bind (fun l -> Some (l.Url.ToString()))
                    |> Option.defaultValue (sprintf "Dangling workflow found. List with id '%s' to which workflow was associated no longer exists. Delete workflow using SharePoint Designer or object model." (id.ToString()))
                | None -> "Not present"

     override this.Visit(wf: Workflow) =
        match wf with
        | BuildIn wf' ->
            this.Lines.Add(sprintf "%s; %s; %s; %s; %s; %s; %s; %A; %s; %s; %s;"
                        "BuildIn"
                        webUrl
                        (wf'.DisplayName.Replace(";", "-").Replace("\n", " "))
                        (wf'.Description.Replace(";", "-").Replace("\n", " "))
                        wf'.LastModifiedBy
                        wf'.RestrictToType
                        (wf'.RestrictToType |> function
                            | "List" -> listIdToString wf'.RestrictToScope
                            | _ as t -> failwithf "Unsupported RestrictType: %s" t)                            
                        wf'.Published
                        (listIdToString wf'.TaskListId)
                        (listIdToString wf'.HistoryListId)
                        "")
        | Nintex wf' ->
            this.Lines.Add(sprintf "%s; %s; %s; %s; %s; %s; %s; %A; %s; %s; %s;"
                        "Nintex"
                        webUrl
                        (wf'.DisplayName.Replace(";", "-").Replace("\n", " "))
                        (wf'.Description.Replace(";", "-").Replace("\n", " "))
                        wf'.LastModifiedBy
                        (wf'.RestrictToType)
                        (wf'.RestrictToType |> function
                            | "List" -> listIdToString wf'.RestrictToScope
                            | _ -> "Unsupported RestrictType")
                        wf'.Published
                        "N/A"
                        "N/A"
                        (match wf'.Region with | Some r -> r | None -> ""))

module WebsWorkflows =
    let generateReport dump =
        let v = WebsWorkflowVisitor()
        DepthFirstTraverser(v).Visit(dump)     
        v.Lines.Insert(0, "WorkflowType; WebUrl; DisplayName; Description; LastModifiedBy; RestrictToType; RestrictToScope; Published; TaskListId (OutOfTheBox); HistoryListId (OutOfTheBox); Region (Nintex)")
        v.Lines.AsReadOnly()