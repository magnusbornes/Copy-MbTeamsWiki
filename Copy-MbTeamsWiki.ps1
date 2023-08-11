<#
.SYNOPSIS
    This script lets you copy a Wiki from one Team Channel to another.
.DESCRIPTION
    This script lets you copy any Teams Wiki from Standard og Private Channels to any other Channel.
    It might work with Shared Channels to, but I never tested that.

    ______________
        How to use:
        1.  Open Teams.
        2.  Right-click, and copy link to the tab in the Channel where the Wiki is located.
                This is your sourceUrl.
        3.  Either, Create a new Teams Wiki tab in the target Channel and copy the link, or copy the link to the Channel (Right-click in left menu!).
                This is your targetUrl. (I recommend doing the former option...)
            
        $SourceUrl = "https://teams.microsoft.com/l/channel/19%3Addb24db9b083441871a99e09148f6dc4%40thread.tacv2/tab%3A%3A95527d65-6cef-4b31-9f68-c21e396444df?groupId=85081890-f758-4db1-b946-a32fe4bec50f&tenantId=0bb2172-822f-4467-a5ba-5bb375967c05"
        
        $TargetUrl = "https://teams.microsoft.com/l/channel/19%3a12fac4bb21db4d6e8e616d0123489a93%40thread.tacv2/GR01-renamed?groupId=85081890-f758-4db1-b946-a32fe4bec50f&tenantId=0bb2172-822f-4467-a5ba-5bb375967c05"
        
        Copy-MbTeamWiki.ps1 -SourceUrl "$"SourceUrl -TargetUrl "$"TargetUrl
        
        This will copy Wiki from the provided source tab, to the provided target tab.
        
    _______________
    As Teams Wiki is being depricated, this script has purose for a limited amout of time.
    Therefore I have not spent a lot of time optimizing the script with parallelization and what not.

    I also only have 2 years of experience with PowerShell, so bear with me.
    The script is pretty messy, but hey; if it works – don't fix it, am I right?

    Created by Magnus Børnes @ Norwegian University of Science and Technology, August 2023.

    "Known" issues:
    -   If image filename already exists, the copy will fail, and you will have to run the script all over, or skip that image.
    -   Some of the regex patterns are pretty jank, and they may fail if you have something unusual in your Wiki, that I haven't tested.
    -   Error handling could absolutely be better, but it's not like any enduser is going to run this script anyway.
    -   Host output could also be better.
.NOTES
    Information or caveats about the function e.g. 'This function is not supported in Linux'
.LINK
    Specify a URI to a help page, this will show when Get-Help -Online is used.
.EXAMPLE
    
    $SourceUrl = "https://teams.microsoft.com/l/channel/19%3Addb24db9b083441871a99e09148f6dc4%40thread.tacv2/tab%3A%3A95527d65-6cef-4b31-9f68-c21e396444df?groupId=85081890-f758-4db1-b946-a32fe4bec50f&tenantId=0bb2172-822f-4467-a5ba-5bb375967c05"
    
    $TargetUrl = "https://teams.microsoft.com/l/channel/19%3a12fac4bb21db4d6e8e616d0123489a93%40thread.tacv2/GR01-renamed?groupId=85081890-f758-4db1-b946-a32fe4bec50f&tenantId=0bb2172-822f-4467-a5ba-5bb375967c05"
    
    Copy-MbTeamWiki.ps1 -SourceUrl "$"SourceUrl -TargetUrl "$"TargetUrl
    
    This will copy Wiki from the provided source tab, to the provided target tab.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)][string]$SourceUrl,
    [Parameter(Mandatory)][string]$TargetUrl,
    [Parameter()][switch]$NoConfirm,
    [Parameter()][switch]$Help,
    [Parameter()][switch]$ExcludeImages
    
)
# Sets PowerShell to Stop if any error occours, to avoid following errors and Armageddon.
$ErrorActionPreference = "Stop"

# Sets base URI and version of Microsoft Graph.
$graphVersion = "v1.0"
$baseUri = "https://graph.microsoft.com/$graphVersion"

Write-Host "`nDisclaimer: Use this script at your own risk. Please test the script on a dummy Wiki with images before running it on something valuable.`n" -f Magenta

# Custom Sleep function which is in use when copying images, and waiting for the copy to complete.
function Start-CustomSleep{
    param(
        [Parameter(Mandatory)][int32]$inputNumber
    )
    $duration = $i*2
    Write-Host "Sleeps for $duration seconds..."
    Start-Sleep $duration
}

# Converts and validates your input-URLs.
function ConvertFrom-TeamUrl {
    [Cmdletbinding()]
    [OutputType([bool])]
	param
	(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
		[string]$InputUrl
	)
    $url = [System.Web.HTTPUtility]::UrlDecode($InputUrl)
    $uri = [System.Uri]($url)
    $uriQuery = [System.Web.HttpUtility]::ParseQueryString($uri.Query)

    $object = [PSCustomObject]@{
        Url = $url
        TenantId = $uriQuery['tenantId']
        TeamId = $uriQuery['groupId']
        ChannelId = $uri.Segments[3].TrimEnd('/')
        TabId = $uri.Segments[4].replace('tab::','')
        UrlType = "Undefined"
    }
    
    # Check if URL is valid, and determine if it links to a channel or a tab.
    if ([guid]::TryParse($object.TeamId, $([ref][guid]::Empty)) -and [guid]::TryParse($object.TabId, $([ref][guid]::Empty))){
        $object.UrlType = "Tab"
    }elseif ([guid]::TryParse($object.TeamId, $([ref][guid]::Empty)) -and ![guid]::TryParse($object.TabId, $([ref][guid]::Empty))){
        $object.UrlType = "Channel"
    }else {
        Write-Error "URL validation error..."
    }

    $object
}

# Import required Modules
if (!(Get-Module Microsoft.Graph.Authentication)) 
    {try{
        Import-Module Microsoft.Graph.Authentication -MinimumVersion 1.0.0; Write-Host "Imports module (Microsoft.Graph.Authentication)..."
    }catch{
        Write-Error $_
        Break
    }
}

# Runs 'ConvertFrom-TeamUrl' to convert input-URLs into objects.
if ($SourceUrl -and $TargetUrl) {
    $source = ConvertFrom-TeamUrl -InputUrl $SourceUrl; Write-Host "Creates source object..."
    $target = ConvertFrom-TeamUrl -InputUrl $TargetUrl; Write-Host "Creates target object..."
}else{
    Write-Error "Missing, or could not validate input URLs..."
    Break
}

# URL-validation: 'SourceUrl' must be a direct link to a Wiki tab.
if($source.UrlType -ne 'Tab'){
    Write-Error "Source URL is not a valid link to a Tab."
    Write-Error "SourceURL = $($source.UrlType)"
    Break
}

# Get Wiki Tab object. Fails if not connected. Run Connect-MgGraph first.
$requestUri = "$baseUri/teams/$($source.teamId)/channels/$($source.channelId)/tabs/$($source.tabId)"
try {
    $wikiTab = Invoke-MgGraphRequest -Uri $requestUri; Write-Host "GET: Teams Wiki Tab object"
}
catch {
    Write-Error $_
    Write-Error "Probably failed to connect to Microsoft Graph. Try 'Connect-MgGraph'."
    Break
}

# Validation: Wiki Tab object must contain a tabID.
# Add WikiTab Id (SharePoint List) and Wiki Tab DisplayName to source object.
Write-Verbose "Validates if Wiki Tab object contains tabID..."
if (!$wikiTab.configuration.wikiTabId) { Write-Error "No wikiTabId found for the specified sourceUrl..."; Break}
Add-Member -InputObject $source -MemberType NoteProperty -Name 'WikiTabId' -Value $($wikiTab.configuration.wikiTabId)
Add-Member -InputObject $source -MemberType NoteProperty -Name 'WikiTabName' -Value $($wikiTab.displayName)

# GET Wiki-tab on target channel, or Create new Tab if input-TargetUrl links to a Channel.

# If 'TargetUrl' is a team/channel link, and not a tab link –> Create a new Wiki Tab in the target channel.
# If 'TargetUrl' is a tab link –> Get target Wiki Tab object.
Write-Host "Validate 'TargetUrl'..."
if ($target.UrlType -ne 'Tab') {
    Write-Warning "TargetUrl is not a tab link. Creating a new tab in the target channel..."
    if (!$NoConfirm) {
        Write-Host "Press 'y' to continue..."
        $continue = Read-Host; if($continue -ne 'y'){Break}
    }
    $newWikiTabRequestBody = @{
        "displayname" = $source.WikiTabName;
        "configuration" = $null;
        "teamsApp@odata.bind" = "$baseUri/appCatalogs/teamsApps/com.microsoft.teamspace.tab.wiki"
    }

    $newWikiTab = Invoke-MgGraphRequest -Method POST -Uri "$baseUri/teams/$($target.teamId)/channels/$($target.channelId)/tabs" -Body $newWikiTabRequestBody
    if($newWikiTab.displayName){
        Write-Host "New Wiki Tab '$($newWikiTab.displayName)' created!" -ForegroundColor Green
        Write-Host "Please open the new tab in Teams and click ""set up tab"", update targetUrl with link to the tab (Right-click the Tab and click 'Copy link to tab'), then Run script again..."
        Break
    }else{
        Write-Error "Something went wrong creating a new Wiki tab :("
        Break
    }
}elseif ($target.UrlType -eq 'Tab') {
    $requestUri = "$baseUri/teams/$($target.TeamId)/channels/$($target.ChannelId)/tabs/$($target.TabId)"
    $newWikiTab = Invoke-MgGraphRequest -Uri $requestUri
    if($newWikiTab.displayName){
        Write-Host "Target Wiki Tab '$($newWikiTab.displayName)' found!" -ForegroundColor Green
    }else{
        Write-Error "Something went wrong getting the targe Wiki tab :("
        Break
    }
}

# Add WikiTab (SharePoint List) Id and Wiki Tab DisplayName to target-object.
Add-Member -InputObject $target -MemberType NoteProperty -Name 'WikiTabId' -Value $($newWikiTab.configuration.wikiTabId)
Add-Member -InputObject $target -MemberType NoteProperty -Name 'WikiTabName' -Value $($newWikiTab.displayName)


# ___________________
# ___ S O U R C E ___
# Gets all the required information we need from the source Wiki, Tab, Channel, Team, Site and List.

# 1.    Detect whether source Wiki is in a Standard or Private channel...
#       (This REST endpoint is only available in 'beta' by August 2023.)
$requestUri = "https://graph.microsoft.com/beta/teams/$($source.teamId)/channels/$($source.channelId)"
$result = Invoke-MgGraphRequest -Uri $requestUri
Add-Member -InputObject $source -MemberType NoteProperty -Name 'membershipType' -Value $result.membershipType
#Add-Member -InputObject $source -MemberType NoteProperty -Name 'ChannelSiteUrl' -Value $((($result.replace('.*\/sites/','').split('/')[0..4] + "/" -join '/').trimend('/'))

Write-Host "GET: SiteId, ListId, and List content of source team..."

# 2.    Get 'SiteId' of source team.
Write-Host "Source channel membershipType: $($source.membershipType)" -f Magenta

if($source.membershipType -eq "private"){
    $splittedChannelUrl = $result.filesFolderWebUrl.replace('.*\/sites/','').split('/')
    $joinedChannelUrl = $splittedChannelUrl[2] + ":/" + ($splittedChannelUrl[3..4] -join '/')
    $requestUri = "$baseUri/sites/$joinedChannelUrl"
}else{
    $requestUri = "$baseUri/groups/$($source.TeamId)/sites/root?`$select=id"
}
Add-Member -InputObject $source -MemberType NoteProperty -Name 'SiteId' -Value $((Invoke-MgGraphRequest -Uri $requestUri).Id.split(',')[1])

# 3.    Get 'ListId' of source Wiki list from SharePoint.
$requestUri = $baseUri + "/sites/$($source.SiteId)/lists/?`$filter=displayName eq '$($source.ChannelId)_wiki'&`$select=id"
Add-Member -InputObject $source -MemberType NoteProperty -Name 'ListId' -Value $((Invoke-MgGraphRequest -Uri $requestUri).Value.Id)

# 4.    Get source Wiki list content from SharePoint.
$wikiFields = "Id,Title,wikiCanvasId,wikiContent,wikiConversationId,wikiDeleted,wikiMetadata,wikiOrder,wikiPageId,wikiSession,wikiSpare1,wikiSpare2,wikiSpare3,wikiTimestamp,wikiTitle,wikiUser"
$requestUri = "$baseUri/sites/$($source.SiteId)/lists/$($source.ListId)/items/?`$expand=fields(`$select=$wikiFields)"
$wikiContent = (Invoke-MgGraphRequest -Uri $requestUri).Value.Fields

# 5.    Get source 'Teams Wiki Data' folder.
$requestUri = "$baseUri/sites/$($source.SiteId)/drives"
$sourceWikiDataDrive = ((Invoke-MgGraphRequest -Uri $requestUri).Value.where{$_.name -eq "Teams Wiki Data"})


# ___________________
# ___ T A R G E T ___
# Gets all the required information we need from the target Wiki, Tab, Channel, Team, Site and List.

# Get SiteId of target team.
Write-Host "GET: SiteId, ListId, and List items of target team..."
$requestUri = "$baseUri/groups/$($target.TeamId)/sites/root?`$select=id"
Add-Member -InputObject $target -MemberType NoteProperty -Name 'SiteId' -Value $((Invoke-MgGraphRequest -Uri $requestUri).Id.split(',')[1])

# Get list Id of target Wiki list from SharePoint.
$requestUri = $baseUri + "/sites/$($target.SiteID)/lists/?`$filter=displayName eq '$($target.ChannelId)_wiki'&`$select=id"
Add-Member -InputObject $target -MemberType NoteProperty -Name 'ListId' -Value $((Invoke-MgGraphRequest -Uri $requestUri).Value.Id)

# Get target 'Teams Wiki Data' folder.
$requestUri = "$baseUri/sites/$($target.SiteId)/drives"
$targetWikiDataDrive = ((Invoke-MgGraphRequest -Uri $requestUri).Value.where{$_.name -eq "Teams Wiki Data"})

# Get target Channel's DisplayName.
$requestUri = "$baseUri/teams/$($target.teamId)/channels/$($target.channelId)"
Add-Member -InputObject $target -MemberType NoteProperty -Name 'ChannelDisplayName' -Value $((Invoke-MgGraphRequest -Uri $requestUri).DisplayName)


# ___________________
# Filter wikiContent to only include the Wiki we want to copy, and not other Wikis in the same Team/Channel.
Write-Verbose "Filters WikiContent to only include the Wiki we want to move, as all Wikis in the chananel is stored in the same list..."
$wikiToCopy = $WikiContent.where{($_.wikiCanvasId -eq $source.wikiTabId) -or ($_.id -eq $source.wikiTabId)}


# ___________________
# For each Page/Section in Wiki:
# A. Checks for Images.
# B. Copies Images to the target Site, and updates HTML Image References for every Image in every Page.
# C. Write HTML-content to the target Wiki.
# D. Update/Patch the Wiki so everything appears in the correct order, relative to the new ListItemIds.

foreach($line in $wikiToCopy){

    # Copy images from source to target.
    if (!$ExcludeImages){ 
        # Section A: Check for Images.
        $images = ([regex]'(<img.*?>)').Matches($line.wikiContent)
        if(!$images){
            Write-Verbose "No images found. Skipping copy-section of script..."
        }else{
            # Checks that there is only 1 folder that matches the target Channel DisplayName.
            $requestUri = "$baseuri/drives/$($targetWikiDataDrive.id)/root/children?`$select=id,name"
            $targetWikiDataFolder = (Invoke-MgGraphRequest -Uri $requestUri).Value.where{$_.name -match $target.ChannelDisplayName}

            # If there is more or less than 1 folder, the scripts ask if the user wants to continue without the images, or abort.
            if($targetWikiDataFolder.count -ne 1){
                Write-Error "Found $($targetWikiDataFolder.count) items with name ""$($target.ChannelDisplayName)"", which indicates something is not quite right..."
                Write-Host "Continue without images? If not, script will be aborted."
                $continue = Read-Host; if($continue -ne 'y'){Break}
                $skipImages = $true
            }

            if(!$skipImages){
                # Section B: Copy Images to the target Site.
                # For every image in the Page:
                foreach($img in $images){

                    # -1. Creates a metadata-object for the image, that will be used for both copying, and updating the HTML Image References later.
                    $imgData = @{
                        TagValue = $img.Value
                        DataSrc = ([regex]'data-src="(.*?)"').Matches($img.Value).Groups[1].Value
                        DataPreviewSrc = ([regex]'data-preview-src="(.*?)"').Matches($img.Value).Groups[1].Value
                        Id = ([regex]'getfilebyid\(%27(.*?)%27\)').Matches($img.Value).Groups[1].Value
                        Name = ([regex]'\/(img.*?.\w*)"').Matches($img.Value).Groups[1].Value
                    }
                    
                    # 0. Prepares payload for POST-request (Copying).
                    $params = @{
                        parentReference = @{
                            driveId = $targetWikiDataDrive.ID
                            id = $targetWikiDataFolder.Id
                        }
                    }

                    # 1.  Run Copy
                    $requestUri = "$baseUri/sites/$($source.SiteId)/drives/$($sourceWikiDataDrive.Id)/items/$($imgData.Id)/copy"
                    function Start-Copy{
                        Clear-Variable newImg
                        Invoke-MgGraphRequest -Method POST -Uri $requestUri -ContentType "application/json; charset=utf-8" -Body $($params | ConvertTo-Json) -ResponseHeadersVariable newImg
                        Return $newImg
                    }
                    $newImg = Start-Copy

                    # 2.  Checks status for the copy, and waits if it's not done.
                    #     User gets some options if copy is taking longer than expected.
                    $i = 0
                    do {
                        if($i -in 0..5){
                            Write-Host "Waiting for Copy-job to complete... $($result.percentageComplete)%..."
                            Start-CustomSleep $i
                        }elseif($i -eq 6){
                            Write-Warning "The copy-job is still not reporting ""Completed"". What to do?"
                            Write-Host "File: $($imgData.Name)"
                            Write-Host "[1] I can see the file in the target folder. Please continue." -f Yellow
                            Write-Host "[2] I can NOT see the file. Please RETRY the copy." -f Yellow
                            Write-Host "[3] I don't care. Skip it. " -f Yellow
                            $read = Read-Host
                            switch ($read) {
                                1 {Write-Host "Continuing..."}
                                2 {Write-Host "Retrying..." Start-Copy; $i = 0}
                                3 {Write-Host "Skipping..."; $skipImg = $true}
                            }
                        }
                        Write-Verbose "Getting Copy-status for $($imgData.name)..."
                        $result = Invoke-MgGraphRequest -Uri $($newImg.Location)
                        $i++
                    } until (
                        ($result.status -eq "completed") -or ($i -gt 6)
                    )
                    
                    if(!$skipImg){

                        # 3.  GET SharePoint Ids for the copied Image.
                        $driveId = $targetWikiDataDrive.Id # Ny drive: (b!XFnDPuzp1kKCUAKnjz-aUOTSGL9StNNIsaLplRfHyOWF3G8brF82TL9NG1PYRl2I)
                        $itemId = $result.resourceId #(014AGBOYZOR7IRXDNAV5AI2VMLCMJOROHN)
                        $requestUri = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/?`$select=webUrl,sharepointIds,name"
                        $newImgSharePointIds = Invoke-MgGraphRequest -Uri $requestUri
                            

                        # 4.  Created metadata-object for the new Image.
                        $newImage = @{
                            DataSrc = "$($newImgSharePointIds.sharepointIds.siteUrl)/_api/web/getfilebyid(%27$($newImgSharePointIds.sharepointIds.listItemUniqueId)%27)/`$value"
                            DataPreviewSrc = [System.Web.HttpUtility]::UrlDecode(($newImgSharePointIds.webUrl).split('.sharepoint.com')[1])
                            Name = $newImgSharePointIds.name
                            Id = $newImgSharePointIds.sharepointIds.listItemUniqueId
                        }
                        # 5.  Generates a new <img>-tag with the updated Image References..
                        $newValue = $imgData.TagValue
                        $newValue = $newValue.replace($imgData.DataSrc,$newImage.DataSrc).replace($imgData.DataPreviewSrc,$newImage.DataPreviewSrc).replace($imgData.Name,$newImage.Name).replace($imgData.Id,$newImage.Id)

                        # 6.  Updates the SharePoint-list (where the Wiki Content is stored) with info from the metadata-object of the copied Image.
                        $line.WikiContent = $line.WikiContent.replace($imgData.TagValue,$newValue)

                    }
                } 
            }
        }
    }

    #   Section C: Write HTML-content to the target Wiki.
    #   Prepares payload for POST-request to write WikiContent to the target Wiki.

    #   Prepare POST-request for writing to the target Wiki.
    $params = @{
        fields = @{
            Title = $line.Title
            wikiCanvasId = $target.wikiTabId
            wikiContent = $line.wikiContent
            wikiConversationId = $line.wikiConversationId
            wikiDeleted = $line.wikiDeleted
            wikiMetadata = $line.wikiMetadata
            wikiOrder = $line.wikiOrder # This will be updated later.
            wikiPageId = $line.wikiPageId # This will be updated later.
            wikiSession = $line.wikiSession
            wikiSpare1 = $line.wikiSpare1
            wikiSpare2 = $line.wikiSpare2
            wikiSpare3 = $line.wikiSpare3
            wikiTimestamp = $line.wikiTimestamp
            wikiTitle = $line.wikiTitle
            wikiUser = $line.wikiUser
        }
    }
    
    #   Writes content to the target Wiki.
    Write-Host "POST: Adding item '$($line.Title)' to '$($target.wikiTabName)'..."
    $requestUri = "$baseUri/sites/$($target.SiteID)/lists/$($target.ListId)/items"
    try {
        $newLine = (Invoke-MgGraphRequest -Method POST -Uri $requestUri -ContentType "application/json; charset=utf-8" -Body $($params | ConvertTo-Json))
    }
    catch {
        Write-Error $_
        Write-Host "Continue?"
        $continue = Read-Host
        if ($continue -ne 'y') {
            Write-Host "Aborted." -ForegroundColor Red
            Break
        }
    }
    if ($line.newId) {
        Write-Warning "newId already exists. Will be overwritten."
        $line.newId = $newLine.Id
    }else{
        $line.Add("newId", $($newLine.Id))
    }
    Clear-Variable newLine
}

#   Section D: Update/Patch the Wiki so everything appears in the correct order, relative to the new ListItemIds. 
#   Fields affected: 'wikiOrder' and 'wikiPage'.
#   (This is to make sure everything will appear in the right order.)

# Prepare the PATCH-request:
foreach($line in $wikiToCopy){

    #   Create an array of the wikiOrder, if wikOrder is not empty.
    if($line.wikiOrder){
        $regex = ([regex]'(\d+)').Matches($line.wikiOrder)
        $wikiOrderArray = @($regex.Value)

        # Then, for each number in the wikiOrder array; update the number with the newId of that object.
        for($i = 0; $i -lt $wikiOrderArray.Length; $i++){
            $wikiOrderArray[$i] = ($wikiToCopy.where{$_.id -eq $wikiOrderArray[$i]}).newId
        }
        # Format wikiOrder array back to string. Format: [1,2,3]

        Write-Host "UPDATE: Item '$($line.Id)' - wikiOrder from '$($line.wikiOrder)' to '[$($wikiOrderArray -join ',')]'"
        $line.wikiOrder = "[$($wikiOrderArray -join ',')]"
    }
    # We do the same for pageId, except it's only a integer, not an array.
    if ($line.wikiPageId) {
        Write-Host "UPDATE: Item '$($line.Id)' - wikiPageId from '$($line.wikiPageId)' to '$(($wikiToCopy.where{$_.id -eq $line.wikiPageId}).newId)'"
        $line.wikiPageId = ($wikiToCopy.where{$_.id -eq $line.wikiPageId}).newId
    }
}

# Now for the actual update (PATCH) of the SharePoint list.
foreach($line in $wikiToCopy){
    $params = @{
        wikiOrder = $line.wikiOrder
        wikiPageId = $line.wikiPageId
    }

    Write-Host "PATCH: Item: '$($line.newId)' - Fields: 'wikiOrder'='$($line.wikiOrder)', 'wikiPageId'='$($line.wikiPageId)'"
    $requestUri = "$baseUri/sites/$($target.SiteID)/lists/$($target.ListId)/items/$($line.newId)/fields"
    try {
        $newLine = (Invoke-MgGraphRequest -Method PATCH -Uri $requestUri -Body $($params | ConvertTo-Json))
    }
    catch {
        Write-Error $_
        Write-Host "Continue?"
        $continue = Read-Host
        if ($continue -ne 'y') {
            Write-Host "Aborted." -ForegroundColor Red
            Break
        }
    }
}

#   Phew. That was pretty tenseful...
#   RIP Teams Wiki. Someone will probably miss you. I won't.
Write-Host Done. -ForegroundColor Green