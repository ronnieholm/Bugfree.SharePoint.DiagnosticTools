namespace SPReports

open System
open System.Security
open System.Collections.Generic
open Microsoft.SharePoint.Client
open Microsoft.Online.SharePoint.TenantAdministration

[<AutoOpen>]
module Helpers =
    let createClientContext (username: string) (password: string) (url: string) =
        let securePassword = new SecureString()
        password.ToCharArray() |> Seq.iter securePassword.AppendChar
        new ClientContext(url, Credentials = SharePointOnlineCredentials(username, securePassword))

    let getSiteCollections (createClientContext: string -> ClientContext) includeDetail (adminUrl: string) =
        let rec getSiteCollectionsRecursive (t: Tenant) (p: List<SiteProperties>) startPosition =
            let ctx = t.Context
            let sites = t.GetSiteProperties(startPosition, includeDetail)
            ctx.Load(sites)
            ctx.ExecuteQuery()
            
            sites |> Seq.iter p.Add
            if sites.NextStartIndex = -1 then p else getSiteCollectionsRecursive t p sites.NextStartIndex
        
        let tenantCtx = createClientContext adminUrl |> Tenant
        getSiteCollectionsRecursive tenantCtx (List<SiteProperties>()) 0

    let getByKey (properties: IDictionary<string, string>) key =
        if properties.ContainsKey(key) then properties.[key]
        else failwithf "Key not found: %s" key

    let getOptionByKey (properties: IDictionary<string, string>) key =
        if properties.ContainsKey(key) then Some(properties.[key])
        else None

    let getAbsoluteListUrl (l: List): Uri =
        let tokens = l.DefaultViewUrl.Split([| '/' |])
        let length = tokens.Length
        let isDocumentLibrary = tokens.[length - 2] = "Forms"

        tokens
        |> Seq.take (if isDocumentLibrary then (length - 2) else (length - 1)) 
        |> Seq.skip 3
        |> Seq.fold (sprintf "%s/%s") l.Context.Url |> Uri