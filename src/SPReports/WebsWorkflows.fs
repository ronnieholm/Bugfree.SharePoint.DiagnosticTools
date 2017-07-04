namespace SPReports

// Loosely inspired by
//   https://community.nintex.com/community/build-your-own/nintex-for-office-365/blog/2017/02/20/migrate-your-tenant-to-another-datacenter-1-of-3
//   https://community.nintex.com/community/build-your-own/nintex-for-office-365/blog/2017/02/22/migrate-your-tenant-to-another-datacenter-2-of-3-or-where-is-the-workflow-app
//   https://community.nintex.com/community/build-your-own/nintex-for-office-365/blog/2017/04/20/migrate-your-tenant-to-another-datacenter-3-of-3-where-are-all-the-workflows

open System
open System.Collections.Generic
open SPReports.DumpMetadata

module WebsWorkflows =
    let lines = ResizeArray<string>()
    let guidListUrl = Dictionary<Guid, string>()

    let rec visitWeb level (w: SPWeb) =        
        printfn "W: %s%s" (String.replicate level " ") w.Url               
        let lookupListByGuid g = 
            match w.Lists with
            | Some lists -> lists |> Array.tryFind (fun l -> l.Id = g)
            | None -> None

        let listIdToString id =
            match id with
            | Some id ->
                lookupListByGuid id
                |> Option.bind (fun l -> Some (l.Url.ToString()))
                |> Option.defaultValue (sprintf "Dangling workflow found. List with id '%s' to which workflow was associated no longer exists." (id.ToString()))
            | None -> "Not present"

        match w.Workflows with
        | Some workflows ->
            for wf in workflows do
                match wf with
                | BuildIn wf' ->
                    lines.Add(sprintf "%s; %s; %s; %s; %s; %s; %s; %s; %A; %s; %s; %s;"
                                "BuildIn"
                                w.Url
                                (wf'.DisplayName.Replace(";", "-").Replace("\n", " "))
                                (wf'.Description.Replace(";", "-").Replace("\n", " "))
                                wf'.LastModifiedBy
                                (wf'.LastModifiedAt.ToUniversalTime().ToString())
                                wf'.RestrictToType
                                (match wf'.RestrictToType with
                                 | "List" -> listIdToString wf'.RestrictToScope
                                 | _ as t -> failwithf "Unsupported RestrictType: %s" t)                            
                                wf'.Published
                                (listIdToString wf'.TaskListId)
                                (listIdToString wf'.HistoryListId)
                                "")
                | Nintex wf' ->
                    lines.Add(sprintf "%s; %s; %s; %s; %s; %s; %s; %s; %A; %s; %s; %s;"
                                "Nintex"
                                w.Url
                                (wf'.DisplayName.Replace(";", "-").Replace("\n", " "))
                                (wf'.Description.Replace(";", "-").Replace("\n", " "))
                                wf'.LastModifiedBy
                                (wf'.LastModifiedAt.ToUniversalTime().ToString())
                                wf'.RestrictToType
                                (match wf'.RestrictToType with
                                 | "List" -> listIdToString wf'.RestrictToScope
                                 | _ -> "Unsupported RestrictType")
                                wf'.Published
                                "N/A"
                                "N/A"
                                (match wf'.Region with | Some r -> r | None -> ""))
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
        [| yield "WorkflowType; WebUrl; DisplayName; Description; LastModifiedBy; LastModifiedAt; RestrictToType; RestrictToScope; Published; TaskListId (OutOfTheBox); HistoryListId (OutOfTheBox); Region (Nintex)"
           for l in lines do yield l |]