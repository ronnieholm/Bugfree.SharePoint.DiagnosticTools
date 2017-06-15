module Program

open System
open System.IO
open System.Text
open System.Diagnostics
open Argu
open MBrace.FsPickler
open SPReports
open SPReports.DumpMetadata
open SPReports.WebsFileExtensionsCount
open SPReports.WebsAddIns
open SPReports.WebsWorkflows

type ListSolutionsFromTenantArgs =
    | [<Mandatory>] Username of username: string
    | [<Mandatory>] Password of password: string
    | [<Mandatory>] Tenant_name of name: string
    | [<Mandatory>] Csv_Output_path of path: string
with
    interface IArgParserTemplate with
        member s.Usage =
            match s with
            | Username _ -> "username of the form <user>@<tenant>.onmicrosoft.com."
            | Password _ -> "password matching username."
            | Tenant_name _ -> "tenant name part of the admin center url of the form https://<tenant>-admin.sharepoint.com."
            | Csv_Output_path _ -> "path and filename to CSV file location."

type DumpMetadataArgs =
    | [<Mandatory>] Username of username: string
    | [<Mandatory>] Password of password: string
    | [<Mandatory>] Tenant_name of name: string
    | [<Mandatory>] Output_path of path: string
with
    interface IArgParserTemplate with
        member s.Usage =
            match s with
            | Username _ -> "username of the form <user>@<tenant>.onmicrosoft.com."
            | Password _ -> "password matching username."
            | Tenant_name _ -> "tenant name part of the admin center url of the form https://<tenant>-admin.sharepoint.com."
            | Output_path _ -> "path to binary dump."

type WebsFileExtensionsCountArgs =
    | [<Mandatory>] Input_path of path: string
    | [<Mandatory>] Output_path of path: string
with
    interface IArgParserTemplate with
        member s.Usage =
            match s with
            | Input_path _ -> "path to binary dump."
            | Output_path _ -> "path to CSV file."

type WebsAddInsArgs =
    | [<Mandatory>] Input_path of path: string
    | [<Mandatory>] Output_path of path: string
with
    interface IArgParserTemplate with
        member s.Usage =
            match s with
            | Input_path _ -> "path to binary dump."
            | Output_path _ -> "path to CSV file."   

type WebsWorkflowsArgs =
    | [<Mandatory>] Input_path of path: string
    | [<Mandatory>] Output_path of path: string
with
    interface IArgParserTemplate with
        member s.Usage =
            match s with
            | Input_path _ -> "path to binary dump."
            | Output_path _ -> "path to CSV file."   

type CLIArguments =
    | Dump_metadata of ParseResults<DumpMetadataArgs>
    | Webs_file_extensions_count of ParseResults<WebsFileExtensionsCountArgs>
    | Webs_add_ins of ParseResults<WebsAddInsArgs>
    | Webs_Workflows of ParseResults<WebsWorkflowsArgs>
    | List_solutions_from_tenant of ParseResults<ListSolutionsFromTenantArgs>
with
    interface IArgParserTemplate with
        member this.Usage =
            match this with
            | Dump_metadata _ -> "dump metadata."
            | Webs_file_extensions_count _ -> "analyze dump and summarize files by extension and count."
            | Webs_add_ins _ -> "analyse dump and summarize webs by add ins."
            | Webs_Workflows _ -> "..."
            | List_solutions_from_tenant _ -> "lists solutions from tenant."

let parsedArgsOrException args =
    try
        let parser = ArgumentParser.Create<CLIArguments>()
        Choice1Of2 (parser.Parse(args))
    with
    | ex -> Choice2Of2 ex

[<EntryPoint>]
let main argv =
    match parsedArgsOrException argv with
    | Choice1Of2 args ->
        match args.TryGetSubCommand() with
        | Some cmd ->
            match cmd with
            | Dump_metadata r ->
                let sw = Stopwatch()
                sw.Start()

                let dump = DumpMetadata.dump (r.GetResult <@ DumpMetadataArgs.Username @>) (r.GetResult <@ DumpMetadataArgs.Password @>) (r.GetResult <@ DumpMetadataArgs.Tenant_name @>)
                let binary = FsPickler.CreateBinarySerializer()
                let pickle = binary.Pickle dump
                File.WriteAllBytes((r.GetResult <@ DumpMetadataArgs.Output_path @>), pickle)
                
                sw.Stop()
                printfn "Finished at %A" (sw.Elapsed)

            | Webs_file_extensions_count r ->
                let binary = FsPickler.CreateBinarySerializer()
                let dump = binary.UnPickle<SPSiteCollection[]>(File.ReadAllBytes(r.GetResult <@ WebsFileExtensionsCountArgs.Input_path @>))
                let csv = WebsFileExtensionsCount.generateReport dump
                File.WriteAllLines(r.GetResult <@ WebsFileExtensionsCountArgs.Output_path @>, csv, Encoding.UTF8)

            | Webs_add_ins r ->
                let binary = FsPickler.CreateBinarySerializer()
                let dump = binary.UnPickle<SPSiteCollection[]>(File.ReadAllBytes(r.GetResult <@ WebsAddInsArgs.Input_path @>))
                let csv = WebsAddIns.generateReport dump
                File.WriteAllLines(r.GetResult <@ WebsAddInsArgs.Output_path @>, csv, Encoding.UTF8)

            | Webs_Workflows r ->
                let binary = FsPickler.CreateBinarySerializer()
                let dump = binary.UnPickle<SPSiteCollection[]>(File.ReadAllBytes(r.GetResult <@ WebsWorkflowsArgs.Input_path @>))
                let csv = WebsWorkflows.generateReport dump
                File.WriteAllLines(r.GetResult <@ WebsWorkflowsArgs.Output_path @>, csv, Encoding.UTF8)

            | List_solutions_from_tenant r ->
                let sandboxSolutions = ListSolutionsFromTenant.generateReport (r.GetResult <@ ListSolutionsFromTenantArgs.Username @>) (r.GetResult <@ ListSolutionsFromTenantArgs.Password @>) (r.GetResult <@ ListSolutionsFromTenantArgs.Tenant_name @>)
                let nl = System.Environment.NewLine
                let header = sprintf "SiteCollectionUrl; WSPName; Author; CreatedAt; Actived; HasAssemblies; SolutionHash%s" nl

                let csv = 
                    sandboxSolutions
                    |> Array.fold (fun acc c ->
                        let line = 
                            sprintf "%s; %s; %s; %A; %A; %A; %s" 
                                c.SiteCollectionUrl 
                                c.WspName  
                                c.Author 
                                c.CreatedAt 
                                c.Activated 
                                c.HasAssemblies 
                                c.SolutionHash
                        acc + line + nl) header
                File.WriteAllText(r.GetResult <@ ListSolutionsFromTenantArgs.Csv_Output_path @>, csv, Encoding.UTF8)
                printfn "Wrote %s" (r.GetResult <@ ListSolutionsFromTenantArgs.Csv_Output_path @>)

        | None -> 
            printfn "%s" (args.Parser.PrintUsage())

    | Choice2Of2 ex  ->
        printfn "%s" ex.Message

    0