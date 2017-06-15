namespace SPReports

// Inspired by https://github.com/OfficeDev/PnP-Tools/tree/master/Scripts/SharePoint.Sandbox.ListSolutionsFromTenant

open System
open System.Text.RegularExpressions
open Microsoft.SharePoint.Client
open SPReports.Helpers

module ListSolutionsFromTenant =
    type SandboxSolution =
        { SiteCollectionUrl: string
          WspName: string
          Author: string
          CreatedAt: DateTime
          Activated: bool      
          HasAssemblies: bool
          SolutionHash: string }

    let stringToBool = function
        | "0" -> false
        | "1" -> true
        | _ -> failwith "Unsupported value" 

    let getSolutions (createClientContext: string -> ClientContext) siteCollectionUrl =    
        use ctx = createClientContext siteCollectionUrl
        let query = CamlQuery.CreateAllItemsQuery()
        let catalog = ctx.Web.GetCatalog(121)
        let items = catalog.GetItems(query)
        ctx.Load(items)
        ctx.Load(catalog)
        ctx.ExecuteQuery()

        items 
        |> Seq.cast<ListItem>
        |> Seq.toArray
        |> Array.map (fun li -> 
            let status = li.["Status"] :?> FieldLookupValue                  
            let re = Regex("SolutionHasAssemblies:IW\|(?<hasAssemblies>\d+)\r\n")
            let m = re.Match(li.["MetaInfo"] |> string)
            let g = m.Groups.["hasAssemblies"]

            { SiteCollectionUrl = siteCollectionUrl
              WspName = li.["FileLeafRef"] |> string
              HasAssemblies = stringToBool g.Value
              Activated =  if status = null then false else stringToBool status.LookupValue
              Author = (li.["Author"] :?> FieldLookupValue).LookupValue
              CreatedAt = li.["Created"] :?> DateTime
              SolutionHash = li.["SolutionHash"] |> string })

    let generateReport userName password tenantName =
        let createClientContext' = createClientContext userName password
        getSiteCollections createClientContext' false (sprintf "https://%s-admin.sharepoint.com" tenantName)
        |> Seq.map (fun sc -> sc.Url)
        |> Seq.filter (fun url -> not <| url.Contains("-public.sharepoint.com"))
        |> Seq.toArray
        |> Array.Parallel.map (fun url -> 
            printfn "%s" url
            getSolutions createClientContext' url)
        |> Array.collect id