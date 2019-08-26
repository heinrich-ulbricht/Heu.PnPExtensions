# DEMO SCRIPT that shows how to use the PnPExpressionGenerator.dll

# enter your site URL here
$siteUrl = "https://heinrichulbricht.sharepoint.com/sites/dev"
$credentials = "heinrich" # comment out this line to get interactive login

$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition


# this loads our library
Add-Type -Path "$scriptPath\PnPExpressionGenerator\bin\Debug\PnPExpressionGenerator.dll" -ErrorAction Stop
$pnp = New-Object PnPExtensions.PnPExpressionGenerator

# connect
SharePointPnPPowerShell2013\Connect-PnPOnline $siteUrl -Credentials $credentials

# let's go: get some web properties
$ctx = Get-PnPContext
$web = Get-PnPWeb
$expr = $pnp.GetExpressions($web, "RoleAssignments.Member.LoginName", "RoleAssignments.RoleDefinitionBindings.Name", "RoleAssignments.RoleDefinitionBindings.RoleTypeKind")
$ctx.Load($web, $expr)
Invoke-PnPQuery

$web.RoleAssignments | % { Write-Host $_.RoleDefinitionBindings.Name $_.RoleDefinitionBindings.RoleTypeKind }
$web.RoleAssignments | % { Write-Host $_.Member.LoginName }

# now we want to get some info about lists
$ctx = Get-PnPContext
$expr = $pnp.GetExpressions($web, "Lists.Hidden", "Lists.Title", "Lists.RootFolder.ServerRelativeUrl", "Lists.DefaultView.Title", "Lists.Fields.StaticName")
$ctx.Load($web, $expr)
Invoke-PnPQuery

# output all non-hidden lists with info we got
$web.Lists | ? { !$_.Hidden } | % { Write-Host "Title: '$($_.Title)' [$($_.RootFolder.ServerRelativeUrl)] Default view: '$($_.DefaultView.Title)' First 5 fields: $($_.Fields | Select-Object -First 5 | % { $_.StaticName })"  }


# now get documents from a library, the file path, filtered by HasUniqueRoleAssignments
$list = $web.Lists.GetByTitle("Documents")
$items = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
$expr = $pnp.GetExpression($items, "FileRef")
$filter = $pnp.GetWhereExpression($items, "b => b.HasUniqueRoleAssignments")
$ctx.Load($items, $filter, $expr)

Invoke-PnPQuery

$items.Count
$items[0]["FileRef"]