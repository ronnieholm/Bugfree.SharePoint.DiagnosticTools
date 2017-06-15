#if INTERACTIVE
#I "../../packages/Selenium.WebDriver.3.4.0/lib/net40"
#r "../../packages/canopy.1.4.5/lib/canopy.dll"
#r "../../packages/FSharp.Data.2.3.3/lib/net40/FSharp.Data.dll"
#endif

open canopy
open FSharp.Data

// Adjust to local environment
let frontpage = "https://bugfree.sharepoint.com"
let username = "rh@bugfree.onmicrosoft.com"
let password = "secret"
let websAddInsCsvPath = "path/to/Webs-add-ins.csv"

type addIns = CsvProvider<Separators = ";", HasHeaders = true, Sample = "Webs-add-ins-sample.csv">
let current = addIns.Load websAddInsCsvPath

let urls =
    current.Rows 
    |> Seq.filter (fun r -> 
        (r.Title.Trim() = "Nintex Workflow for Office 365" && r.Version.Trim() <> "1.0.4.0") ||
        (r.Title.Trim() = "Nintex Forms for Office 365" && r.Version.Trim() <> "1.2.3.0"))
    |> Seq.map (fun r -> (r.Id.ToString(), r.WebUrl.Trim()))
    // Update add-ins in batches to avoid overloading the platform
    |> Seq.skip 0
    |> Seq.take 25

#if INTERACTIVE
#else
[<EntryPoint>]
let main argv = 
#endif
    configuration.chromeDir <- (sprintf "%s\..\..\packages\Selenium.WebDriver.ChromeDriver.2.30.0.1\driver\win32" __SOURCE_DIRECTORY__)
    start chrome

    // Browser window is an isolated session. By navigating to a SharePoint page, we trigger username/password prompt
    url frontpage
    "#cred_userid_inputtext" << username
    "#cred_password_inputtext" << password
    // Wait for JavaScript to enable sign-in button
    sleep 1.
    click "#cred_sign_in_button"

    for id, webUrl in urls do
        printfn "%s" webUrl
        url (sprintf "%s/_layouts/15/viewlsts.aspx" webUrl)        
        try
            configuration.elementTimeout <- 5.
            click (sprintf "a[href='javascript:handleUpdateApp(\\'%s\\');']" id)
            // Previous step triggers a in-page pop-up/overlay which can take some time load
            configuration.elementTimeout <- 60.
            click "button[name='trialbutton']"
            // Name has ASP.NET generated value of the form "ctl00$PlaceHolderMain$BtnAllow"
            click "input[name*='BtnAllow']"
            // Don't navigate away from the page before the update has started
            sleep 10.
        with
        | :? CanopyElementNotFoundException -> 
            printfn "Application Id %s on '%s' already either updated or in process of being updated" id webUrl

    0