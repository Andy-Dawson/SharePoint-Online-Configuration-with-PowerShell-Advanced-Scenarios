# This PowerShell script accompanies the presentation "SharePoint Online Configuration with PowerShell – Advanced Scenarios.pptx"

# Updating modules
Update-Module -Name SharePointPnPPowerShellOnline
# Get a list of the versions installed
Get-InstalledModule -Name SharePointPnPPowerShellOnline -AllVersions
# Remove an old version
Get-InstalledModule -Name SharePointPnPPowerShellOnline -RequiredVersion 3.6.1902.2 | Uninstall-Module

# Automated removal of all except the latest version of the module
$LatestModule = Get-InstalledModule -Name SharePointPnPPowerShellOnline
$AllModules = Get-InstalledModule -Name SharePointPnPPowerShellOnline -AllVersions
foreach ($m in $AllModules)
{
    if ($m.version -ne $LatestModule.version)
	{
	    Write-Host "Uninstalling $($m.name) - $($m.version) [latest is $($LatestModule.version)]"
	    $m | uninstall-module -force
    }
}



# Making a connection
# Connect to the the root site collection
Connect-PnPOnline -Url https://mydomain.sharepoint.com
# Connect to a specific site collection
Connect-PnPOnline -Url https://mydomain.sharepoint.com/sites/DDDHull
# Note that credentials are not requested if a set of credentials whose URL matches those of the web application we're attempting to connect to are present in the Windows Credential Manager



# Group-less Team site (collection)
# Note: This should also be avilable in the Modern GUI when Offie 365 Group creation is disabled
New-PnPTenantSite -Title "Group-less Team Site" -Url https://mydomain.sharepoint.com/sites/DDDHullGroupless -Description "Group-less Team Site" -Owner Andy@mydomain.onmicrosoft.com -Template STS#3 -TimeZone 2



# Content types, templates, libraries, views
# An aside first
Connect-PnPOnline -Url https://mydomain.sharepoint.com/sites/DDDHull -CreateDrive
# Generates a drive 'SPO:\'
# cd SPO:\
# dir
# etc.

# Creating fields
# Basic fields using the defaults for SharePoint
Add-PnPField -DisplayName "Policy Owner" -InternalName "PolicyOnwer" -Type User -Group "mydomain Columns"
Add-PnPField -DisplayName "Policy Review Date" -InternalName "PolicyReviewDate" -Type DateTime -Group "mydomain Columns"

# More control is provided by using XML when creating the fields
Add-PnPFieldFromXml '<Field Type="DateTime" DisplayName="Policy Review Date 2" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" Format="DateOnly" Group="mydomain Columns" FriendlyDisplayFormat="Disabled" ID="{7831d5d2-6a2e-462a-96fb-976f61df1df7}" StaticName="PolicyReviewDate2" Name="PolicyReviewDate2"></Field>'

# Sample SchemaXml fields from a couple of created fields that don't use the defaults
# DateTime using DateOnly
#SchemaXml                     : <Field Type="DateTime" Name="PolicyReviewDate" DisplayName="Policy Review Date" ID="{9a8b99a0-7b0f-43b2-a827-ccead0742caa}" Group="mydomain Columns" Required="FALSE" SourceID="{f8e19d01-b026-4360-ac15-44cc4d08138e}" StaticName="PolicyReviewDate" Version="2" CustomFormatter="" EnforceUniqueValues="FALSE" Indexed="FALSE" CalType="0" Format="DateOnly" FriendlyDisplayFormat="Disabled" />
# User field using PeopleOnly
#SchemaXml                        : <Field Type="User" Name="PolicyOnwer" DisplayName="Policy Owner" ID="{8bab870a-debf-408d-a925-2c83fbce6aab}" Group="mydomain Columns" Required="FALSE" SourceID="{f8e19d01-b026-4360-ac15-44cc4d08138e}" StaticName="PolicyOnwer" Version="2" EnforceUniqueValues="FALSE" ShowField="ImnName" UserSelectionMode="PeopleOnly" UserSelectionScope="0" />

# Now add the content type and add the templat
$filepath = "C:\src\MyDocumentTemplate.dotx"
$filename = $filePath.Split("\")[-1]
$ctname = "NewDocument"
 
# Create the content type - specify your own
Add-PnPContentType -Name $ctname -ContentTypeId 0x010100deede662c3a94a74832ef7da5a1b7eee -Group "mydomain Content Types" -Description "Create a new NewDocument."
Add-PnPFieldToContentType -ContentType -Field

# Upload the document template to the corresponding folder in _cts
Add-PnPFile -Path $filepath -Folder "/_cts/$ctname" # No need to pre-create the folder
 
# get the content type
$ct = Get-PnPContentType -Identity $ctname
 
# Set the document template for the content type
$ct.DocumentTemplate = $filename # No need to sepcify the full path, just the file name
$ct.Update($true)
Invoke-PnPQuery

# Adding a content type to a list/library
Add-PnPContentTypeToList -List "Documents" -ContentType "NewDocument" -DefaultContentType
# ... and removing the original default
Remove-PnPContentTypeFromList -List "Documents" -ContentType "Document"

# Adding a library view
Add-PnPView -List "Documents" -Title "My New View" -Fields "Title","Address"



# Getting and removing list items
# Get all items on a list
Get-PnPListItem -List "List1"

# Send to the recycle bin
$ListItems = Get-PnPListItem -List "List1"
foreach ($listItem in $ListItems) {
    Move-PnPListItemToRecycleBin -List "List1" -Identity $ListItem.ID -Confirm:$false # Otherwise you get prompted for each list item...
    }

# Actually remove the items
$ListItems = Get-PnPListItem -List "List1"
foreach ($listItem in $ListItems) {
    Remove-PnPListItem -List "List1" -Identity $ListItem.ID -Confirm:$false # Otherwise you get prompted for each list item...
    } # Can use -Recycle on this command, which sends the item to the recycle bin



# Featured List items
# Note: uses internal field names
# Connect to the root site collection first as that is the site this list is in
# Connect to the the root site collection
Connect-PnPOnline -Url https://mydomain.sharepoint.com
Add-PnPListItem -List "SharePointHomeOrgLinks" -Values @{"Title" = "The BBC"; "Url"="http://www.bbc.co.uk"; "Priority"="30000"; "MobileAppVisible"="1"}
# Repeat the add- command for additional links.


# Navigation
# We can use Add-PnPNavigationNode to add navigation nodes
# Remember we can add multi-level specifying parent (needs to be the node number)
Connect-PnPOnline -Url https://mydomain.sharepoint.com/sites/DDDHull
$node = Add-PnPNavigationNode -Location TopNavigationBar -Title "Top Level" -Url "/sites/Test5/" # Top level node
$node2 = Add-PnPNavigationNode -Location TopNavigationBar -Title "Sub-Level" -Url "/sites/Test5/" -Parent $node.id
$node3 = Add-PnPNavigationNode -Location TopNavigationBar -Title "Sub-Sub-Level" -Url "/sites/Test5/" -Parent $node2.id

# Show the navigation node tree; specify top level, quick launch, or search
Get-PnPNavigationNode -Location TopNavigationBar -tree 

# Enable the MegaMenu on a Communications Site
Connect-PnPOnline -Url https://mydomain.sharepoint.com
$web = Get-PnPWeb
$web.MegaMenuEnabled = $true
$web.Update()
Invoke-PnPQuery



# Property bag config
Get-PnPPropertyBag -Key MyKey
Set-PnPPropertyBagValue -Key MyKey -Value MyValue

Get-PnPPropertyBag # Gets all of the property bag values
# And that's it...

# Or we could do it the hard way (CSOM)...
function ReadSPO-PropertyBags 
{ 
    param ($SiteUrl,$UserName,$Password) 
    try 
    {     
        # Add the Client Assemblies - modify the path to match the location on your machine       
        Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\Microsoft.SharePointOnline.CSOM.16.1.8523.1200\lib\net45\Microsoft.SharePoint.Client.dll" 
        Add-Type -Path "C:\Program Files\PackageManagement\NuGet\Packages\Microsoft.SharePointOnline.CSOM.16.1.8523.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll" 
 
        # Client Context 
        $spoCtx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl) 
        $spoCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $Password)   
        $spoCtx.Credentials = $spoCredentials       
 
        $spoSiteCollection=$spoCtx.Site 
        $spoCtx.Load($spoSiteCollection) 
        $spoRootWeb=$spoSiteCollection.RootWeb 
        $spoCtx.Load($spoRootWeb)         
        $spoAllSiteProperties=$spoRootWeb.AllProperties 
        $spoCtx.Load($spoAllSiteProperties) 
        $spoCtx.ExecuteQuery()                 
        $spoPropertyBagKeys=$spoAllSiteProperties.FieldValues.Keys 

        foreach($spoPropertyBagKey in $spoPropertyBagKeys){ 
            Write-Host "PropertyBag Key: " $spoPropertyBagKey " - PropertyBag Value: " $spoAllSiteProperties[$spoPropertyBagKey] -ForegroundColor Green 
        }         
        $spoCtx.Dispose() 
    } 
    catch [System.Exception] 
    { 
        write-host -f red $_.Exception.ToString()    
    }     
} 
 
# Required Parameters 
$SiteUrl = "https://mydomain.sharepoint.com/sites/DDDHull"
$UserName = "Andy@mydomain.SharePoint.com"
# $Password = Read-Host -Prompt "Enter your password: " -AsSecureString   
$Password=convertto-securestring "My Password" -asplaintext -force 

# Now actually call the function...
ReadSPO-PropertyBags -sSiteUrl $SiteUrl -sUserName $UserName -sPassword $Password

## End of file
