namespace SPReports

open System
open System.Linq.Expressions
open System.Xml.Linq
open System.Collections.Generic
open System.Text.RegularExpressions
open System.Net
open Microsoft.SharePoint.Client
open Microsoft.Online.SharePoint.TenantAdministration
open Microsoft.SharePoint.Client.WorkflowServices
open Helpers

module DumpMetadata =
    type Options =
        { IncludeWebs: bool
          IncludeLists: bool
          IncludeListItems: bool
          IncludeWorkflows: bool
          IncludeAddIns: bool }

    type Expr = 
        static member QuoteS(e: Expression<System.Func<Site, obj>>) = e
        static member QuoteW(e: Expression<System.Func<Web, obj>>) = e
        static member QuoteL(e: Expression<System.Func<List, obj>>) = e
        static member QuoteRS(e: Expression<System.Func<RegionalSettings, obj>>) = e

    [<Measure>] type mb
    type XE = XElement
    type XA = XAttribute
    let xn s = XName.Get s

    type SPVersion = { Major: int; Minor: int }

    type SPDocument =
        { Id: Guid
          ItemId: int
          Path: string
          Size: int<mb>
          ContentTypeId: string
          Version: SPVersion
          CreatedAt: DateTime
          ModifiedAt: DateTime }

    type SPListItem =
        { Id: Guid
          ItemId: int
          Version: SPVersion
          CreatedAt: DateTime
          ModifiedAt: DateTime }

    type SPItem =
        | ListItem of SPListItem
        | Document of SPDocument

    type SPList =
        { Id: Guid
          Url: Uri
          Title: string
          Items: SPItem[] option }

    type AddIn =
        { Id: Guid
          ProductId: Guid
          Title: string
          AppWebFullUrl: string
          AppPrincipalId: string
          Version_: string option }     

    type BuildInWorkflow =
        { DisplayName: string
          Description: string
          LastModifiedBy: string
          RestrictToType: string
          RestrictToScope: Guid option
          Published: bool
          TaskListId: Guid option
          HistoryListId: Guid option }

    type NintexWorkflow =
        { DisplayName: string
          Description: string
          LastModifiedBy: string
          RestrictToType: string

          // Has been observed to be null even for RestrictToType = "List"
          RestrictToScope: Guid option
          Published: bool

          // Specific to Nintex workflows
          AppAuthor: string Option
          Region: string option
          Designer: string }        

    type Workflow =
        | BuildIn of BuildInWorkflow
        | Nintex of NintexWorkflow            

    type RegionalSettings =
        { Lcid: int } 

    type SPWeb =
        { Id: Guid
          Url: string
          Title: string
          LastModifiedDate: DateTime
          Created: DateTime
          RegionalSettings: RegionalSettings
          AddInInstances: AddIn[] Option
          Workflows: Workflow[] Option
          Lists: SPList[] Option
          Webs: SPWeb[] Option }

    type SPSiteCollection =
        { Id: Guid
          Url: string
          StorageUsage: int64<mb>
          
          // Autoscaling site collections have a maximum level of 1,048,576MB = 1TB
          StorageMaximumLevel: int64<mb>
          LastContentModifiedDate: DateTime
          Lcid: int
          Template: string

          // Some in case webs are included in the traversal, otherwise None
          RootWeb: SPWeb option }

    let parseUiVersion (fv: Dictionary<string, obj>) =
        // Items uploaded to a list with versioning disabled gets an implicit _UIVersion = 512 
        // and _UIVersionString = "1.0". On the other hand, if an item is uploaded while versioning 
        // is enabled, and then versioning is disabled, the version number gets stuck at the last
        // version number.
        let v = fv.["_UIVersion"] :?> int        
        { Major = v / 512; Minor = v % 512 }

    let parseItem baseListType (li: ListItem) : SPItem =
        match baseListType with
        | BaseType.DocumentLibrary ->
            Document(
                { Id = li.["UniqueId"] :?> Guid
                  ItemId = li.Id
                  Path = li.["FileRef"] |> string
                  Size = 0<mb>
                  ContentTypeId = (li.FieldValues.["ContentTypeId"] |> string).ToLower()
                  Version = parseUiVersion li.FieldValues
                  CreatedAt = DateTime.Parse(li.["Created_x0020_Date"] |> string).ToUniversalTime()
                  ModifiedAt = li.["Modified"] :?> DateTime })
        | _ -> 
            ListItem(
                { Id = li.["UniqueId"] :?> Guid
                  ItemId = li.Id
                  Version = parseUiVersion li.FieldValues
                  CreatedAt = li.["Created"] :?> DateTime
                  ModifiedAt = li.["Modified"] :?> DateTime })

    let createAllItemsQuery (p: ListItemCollectionPosition) =
        let q = 
            CamlQuery(
                ViewXml = 
                    XE(xn "View", XA(xn "Scope", "RecursiveAll"),
                        XE(xn "ViewFields",
                            XE(xn "FieldRef", XA(xn "Name", "Title")),
                            XE(xn "FieldRef", XA(xn "Name", "ContentTypeId")),
                            XE(xn "FieldRef", XA(xn "Name", "_UIVersion"))),
                        XE(xn "RowLimit", "5000"),
                        XE(xn "Query")).ToString())
        q.ListItemCollectionPosition <- p
        q
   
    let productsVersionParsers : Map<Guid, (AppInstance -> string option)> =
        [ // Nintex Forms for Office 365
          Guid "353e0dc9-57f5-40da-ae3f-380cd5385ab9", fun (i: AppInstance) -> 
            // https://formso365.nintex.com/Default.aspx?{StandardTokens}&Src={Source}&remoteAppUrl=https://formso365.nintex.com&platform=MSSharePoint&SPAppVer=1.2.3.0
            let re = Regex("https://.*SPAppVer=(?<v>.*)$")
            let m = re.Match(i.StartPage)
            if m.Success then Some m.Groups.["v"].Value else None

          // Nintex Workflow for Office 365
          Guid "5d3d5c89-3c4c-4b46-ac2c-86095ea300c7", fun i -> 
            // https://workflowo365.nintex.com/Hub.aspx?{StandardTokens}&AppVersion=1.0.4.0
            let re = Regex("https://.*AppVersion=(?<v>.*)$")
            let m = re.Match(i.StartPage)
            if m.Success then Some m.Groups.["v"].Value else None ] 
        |> Map.ofList

    let parseWorkflows level (web: Web): Workflow[] =       
        printf "%s[Workflows]" (String.replicate level " ")
        let ctx = web.Context
        let manager = WorkflowServicesManager(ctx, web)
        let service = manager.GetWorkflowDeploymentService()
        let definitions = service.EnumerateDefinitions(publishedOnly = false)

        ctx.Load(definitions)
        ctx.Load(web)
        ctx.ExecuteQueryWithRetry()

        let workflows =
            [| for d in definitions |> Seq.filter (fun d -> d.RestrictToType = "List") do
                let p = d.Properties
                match getOptionByKey p "AppAuthor" with
                | Some a when a.Contains("Nintex Workflow for Office 365") ->
                    yield Nintex
                        { DisplayName = d.DisplayName
                          Description = d.Description
                          AppAuthor = getOptionByKey p "AppAuthor"
                          LastModifiedBy = getByKey p "ModifiedBy"
                          RestrictToType = d.RestrictToType
                          RestrictToScope = if d.RestrictToScope = null then None else Some (d.RestrictToScope |> Guid)
                          Published = d.Published
                          Region = getOptionByKey p "NWConfig.Region"
                          Designer = getByKey p "NWConfig.Designer" }
                | None -> 
                    yield BuildIn 
                        { DisplayName = d.DisplayName
                          Description = d.Description
                          LastModifiedBy = getByKey p "ModifiedBy"
                          RestrictToType = d.RestrictToType
                          RestrictToScope = if d.RestrictToScope = null then None else Some (d.RestrictToScope |> Guid)
                          Published = d.Published
                          TaskListId = getOptionByKey p "TaskListId" |> Option.map Guid
                          HistoryListId = getOptionByKey p "HistoryListId" |> Option.map Guid }
                | _ -> eprintfn "Unknown type of workflow: %A" p |]

        printfn " (%d)" (Array.length workflows)
        workflows

    let parseAddIns level (web: Web) =
        printf "%s[AddIns]" (String.replicate level " ")
        let ctx = web.Context
        let appInstances = AppCatalog.GetAppInstances(ctx, web)
        ctx.Load(appInstances)
        ctx.ExecuteQueryWithRetry()

        let instances =
            appInstances |> Seq.map (fun i ->
                { Id = i.Id
                  ProductId = i.ProductId
                  Title = i.Title
                  AppWebFullUrl = i.AppWebFullUrl
                  AppPrincipalId = i.AppPrincipalId

                  // No API seems to exist to extract an addIns' version. Some addIns include
                  // the version as part of i.StartPage, but it's in a addIns specific format.                  
                  Version_ =
                    productsVersionParsers
                    |> Map.tryFind i.ProductId
                    |> Option.bind (fun p -> p i) })
                |> Seq.toArray
        
        printfn " (%d)" (Array.length instances)
        instances

    let parseListItems (list: List) (o: Options) =
        let ctx = list.Context
        let rec parseListInternal (items': ListItemCollection) =
            let parseItem' = parseItem list.BaseType

            let items'' = 
                items' 
                |> Seq.cast<ListItem>
                |> Seq.map parseItem'
                |> Seq.toList
            let p = items'.ListItemCollectionPosition
            if p = null then items''
            else
                let q = createAllItemsQuery p
                let items''' = list.GetItems(q)
                ctx.Load(items''')
                ctx.ExecuteQueryWithRetry()
                items'' @ (parseListInternal items''') 
        
        let query = createAllItemsQuery null
        let items = list.GetItems(query)
        ctx.Load(items)
        ctx.ExecuteQueryWithRetry()        
        (parseListInternal items) |> List.toArray

    let parseLists level (web: Web) (o: Options) =
        let ctx = web.Context
        let lists = web.Lists
        ctx.ExecuteQueryWithRetry()  
        [| for l in lists do
            printf "%s[L] %s" (String.replicate level " ") l.Title
            ctx.Load(l, [| Expr.QuoteL(fun l -> l.DefaultViewUrl :> obj); Expr.QuoteL(fun l -> l.BaseType :> obj) |])
            ctx.ExecuteQueryWithRetry()
            yield { Id = l.Id; Url = getAbsoluteListUrl l; Title = l.Title; 
                    Items = 
                        if o.IncludeListItems then 
                            let items = parseListItems l o
                            printfn " (%d)" (Array.length items)
                            Some items
                        else
                            printfn ""
                            None } |]

    let parseRegionalSettings (web: Web): RegionalSettings =
        let ctx = web.Context
        let rs = web.RegionalSettings
        ctx.Load(rs, Expr.QuoteRS(fun r -> r.LocaleId :> obj))
        ctx.ExecuteQueryWithRetry()
        { Lcid = int rs.LocaleId }

    let rec parseWeb level (web: Web) (o: Options) : SPWeb =
        printfn "%s[W] %s" (String.replicate level " ") web.Url        
        let ctx = web.Context
        ctx.Load(web, [|Expr.QuoteW(fun w -> w.Lists :> obj); Expr.QuoteW(fun w -> w.Webs :> obj)|])                      
        ctx.ExecuteQueryWithRetry()

        let nextLevel = level + 2
        { Id = web.Id
          Url = web.Url
          Title = web.Title
          LastModifiedDate = web.LastItemModifiedDate
          Created = web.Created
          RegionalSettings = parseRegionalSettings web
          AddInInstances = if o.IncludeAddIns then Some (parseAddIns nextLevel web) else None
          Workflows = if o.IncludeWorkflows then Some (parseWorkflows nextLevel web) else None
          Lists = if o.IncludeLists then Some (parseLists nextLevel web o) else None
          Webs = if o.IncludeWebs then Some [| for w in web.Webs do yield parseWeb nextLevel w o |] else None }

    let parseSiteCollection (createClientContext: string -> ClientContext) (p: SiteProperties) (o: Options) =
        printfn "[SC] %s" p.Url
        use ctx = createClientContext p.Url
        let site = ctx.Site
        let rootWeb = ctx.Site.RootWeb
        ctx.Load(site, [|Expr.QuoteS(fun s -> s.Id :> obj); Expr.QuoteS(fun s -> s.Url :> obj); Expr.QuoteS(fun s -> s.RootWeb :> obj)|])
        ctx.Load(rootWeb, [|Expr.QuoteW(fun w -> w.Url :> obj); Expr.QuoteW(fun w -> w.Lists :> obj); Expr.QuoteW(fun w -> w.Webs :> obj)|])       
        ctx.ExecuteQueryWithRetry()

        { Id = site.Id
          Url = site.Url
          StorageUsage = p.StorageUsage |> LanguagePrimitives.Int64WithMeasure
          StorageMaximumLevel = p.StorageMaximumLevel |> LanguagePrimitives.Int64WithMeasure
          LastContentModifiedDate = p.LastContentModifiedDate
          Lcid = p.Lcid |> int
          Template = p.Template
          RootWeb = if o.IncludeWebs then Some (parseWeb 0 rootWeb o) else None }

    let dump userName password tenantName =
        //let options = { IncludeWebs = true; IncludeLists = true; IncludeListItems = true; IncludeWorkflows = true; IncludeAddIns = true }
        let options = { IncludeWebs = true; IncludeLists = true; IncludeListItems = false; IncludeWorkflows = true; IncludeAddIns = true }

        let createClientContext' = createClientContext userName password               
        getSiteCollections createClientContext' true (*false*) (sprintf "https://%s-admin.sharepoint.com" tenantName)
        |> Seq.toArray
        |> Array.filter (fun properties -> not <| (properties.Url.ToLower() = (sprintf "https://%s.sharepoint.com/" (tenantName.ToLower())) || properties.Url.Contains("-public.sharepoint.com") || properties.Url.Contains("-my.sharepoint.com")))
        |> Array.mapi (fun i properties -> 
            printfn "%d -- %s" i properties.Url
            try
                Some (parseSiteCollection createClientContext' properties options)
            with    
            | :? WebException as e when e.Message = "The remote server returned an error: (401) Unauthorized." ->
                eprintfn "%s: %s" (properties.Url) e.Message
                None
            | _ -> reraise())
        |> Array.choose id