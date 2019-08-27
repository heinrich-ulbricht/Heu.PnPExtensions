# Heu.PnPExtensions

Currently this consists of PnPExpressionGenerator, a class that allows making complex CSOM queries without losing the simplicity of PnP PowerShell.

## What Problem Does it Solve?

This project allows you to easily retrieve objects and properties from SharePoint via CSOM, even if they are nested deeply. The `PnPExpressionGenerator` class creates the corresponding Lambda expression needed for `ctx.Load(...)`.

The project makes working with CSOM easier especially if you are used to the simplicity of PnP PowerShell and have to resort to CSOM for advanced queries.

Here are examples of Lambda expressions you'd normally use to specify which CSOM properties to load and the corresponding string you can now use instead:

| Base object |  Lambda you might want to use | Property string to generate this lambda |
|------|--------|----
| Web |  `a => a.Title`  | "Title"
| Web |  `a => a["PropertyName"]` | "PropertyName"
| Web |  `a => a.Lists.Include(b => b)` | "Lists"
| Web |  `a => a.Lists.Include(b => b.Title)` | "Lists.Title"
| Web |  `a => a.Lists.Include(b => b.RoleAssignments.Include(c => c))` | "Lists.RoleAssignments"
| Web |  `a => a.Lists.Include(b => b.RoleAssignments.Include(c => c.Member))` | "Lists.RoleAssignments.Member"
| Web |  `a => a.RoleAssignments.Include(b => b.RoleDefinitionBindings.Include(c => c.Name))` | "RoleAssignments.RoleDefinitionBindings.Name"
| Web |  `a => a.Lists.Include(b => b.DefaultView.Title)` | "Lists.DefaultView.Title"
| ListItemCollection |  `a => a.Include(b => b.DisplayName)` | "DisplayName"
| ListItemCollection |  `a => a.Include(b => b["FileRef"])` | "FileRef"

Have a look at `SampleUsage.ps1` to see some of those in action.

## Sample Use Case

Say you want to:
* retrieve all items from a list
* but only those which have permission inheritance broken
* get information about these list item's role assignments
* including member name and role name

You could do this, using mainly PnP PowerShell and loops to retrieve properties one by one:

```
Connect-PnPOnline https://contoso.sharepoint.com/sites/dev
$list = Get-PnPList "SiteAssets"


$listItems = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
$ctx = Get-PnPContext
$ctx.Load($listItems)
Invoke-PnPQuery
foreach ($item in $listItems)
{
    $uniquePermissions = Get-PnPProperty $item "HasUniqueRoleAssignments"
    if (!$uniquePermissions)
    {
        $roleAssignments = Get-PnPProperty $item "RoleAssignments"
        foreach ($ra in $roleAssignments) 
        {
            Get-PnPProperty $ra "RoleDefinitionBindings", "Member"
            $roleDefinitionBindings = $ra.RoleDefinitionBindings
            $member = $ra.Member
            foreach ($rdb in $roleDefinitionBindings)
            {
                Write-Host "$($item[""FileRef""]) - $($member.Title) - $($rdb.Name)"
            }
        }        
    }
}
```
This solution
* retrieves all list items (although only some are needed)
* makes a lot of calls to SharePoint to retrieve the nested properties
* transmits more property values than we actually need
* is a nasty for-loop pyramid


How about this instead:

```
Add-Type -Path "C:\path\to\PnPExpressionGenerator.dll"
$pnp = New-Object PnPExtensions.PnPExpressionGenerator

$listItems = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
$ctx = Get-PnPContext

$filter = $pnp.GetWhereExpression($listItems, "a => !a.HasUniqueRoleAssignments")
$exp1 = $pnp.GetExpression($listItems, "FileRef")
$exp2 = $pnp.GetExpression($listItems, "RoleAssignments.RoleDefinitionBindings.Name")
$exp3 = $pnp.GetExpression($listItems, "RoleAssignments.Member.Title")

$ctx.Load($listItems, $filter, $exp1, $exp2, $exp3)
Invoke-PnPQuery

$listItems | % { $item = $_; $item.RoleAssignments | % { $ra = $_; $member = $ra.Member; $ra.RoleDefinitionBindings | % { $rdb = $_; Write-Host "$($item[""FileRef""]) - $($member.Title) - $($rdb.Name)" }}}
```

This does:
* retrieve list elements that have permission inheritance broken
* along with chosen property values across the object hierarchy
* transmitting only parameter values that we care about
* in one call


## How Does it Work?

Expressions that can be used with `ctx.Load(...)` are generated like this:
1. the text you specify ("_Lists.Title_") is broken in to parts
1. reflection is used to determine whether those parts correspond to a collection property ("_Lists_") or not ("_Title_")
1. having this information a "Lambda string" is generated ("_a => a.Lists.Include(b => b.Title)_")
1. the "Lambda string" is put into the .NET compiler platform (a.k.a. Roslyn) to generate the actual expression code

The PnPExpressionGenerator has no direct dependency on the SharePoint Client Library (`Microsoft.SharePoint.*.dll`). This makes it independent from any specific version of the SP Client Library. Thus once build it should work with any version without needing to rebuild against this specific version.

## Prerequisites
* .NET Framework 4.6.1 (note that you usually don't have this on older on-prem systems)

## How to Build and Use

1. clone the repository
2. build the project and copy the output somewhere
3. load the `PnPExpressionGenerator.dll` into your PowerShell session
4. generate expressions

Look at `SampleUsage.ps1` for an example of how to do this.

Binary builds might follow.
