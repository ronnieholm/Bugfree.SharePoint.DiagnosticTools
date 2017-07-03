$base = Split-Path $PSCommandPath
$msbuild = "C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin\amd64\msbuild.exe"
$sln = "Bugfree.SharePoint.DiagnosticTools"
&$msbuild "$($base)\$($sln).sln"

$date = Get-Date
$date = $date.ToString("yyyyMMdd")
$release = "$($base)\Releases\$($sln)-$($date)"

if (Test-Path $release) {
    Remove-Item -Recurse -Force $release
}
New-Item -Type Directory -Path $release

@("SPAccessMethodLatency",
  "SPDatabaseInspector",
  "SPReports",
  "SPSearchIndexLatency",
  "SPTransferSpeed",
  "SPWorkflowLatency") | foreach { 
	Copy-Item "$($base)\src\$($_)\bin\Debug\*" $release 
}
