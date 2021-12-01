## European Collaboration Conference 2021 Teams Graph API 1.12.2021
##
##  How Teams Graph API can be used to create, manage and manipulate teams
##   
##  Vesa Nopanen
##  Modern and Future Work
##  Principal Consultant, Microsoft MVP
##  Sulava Oy
##
##

# Agenda
# Create team using Teams team template
# Get Join link to the team
# Update team settings
# Provisioning email to General channel
# Update team photo
# Add applications
# Remove Wikis
# Copy template documents to the team
# Attach copied word document to channel tab
# Attach Document library to channel tab
# Add existing tenant guest to the team
# Create a tag to team
# Update channel moderation settings
# Create private channel
# Create and attach OneNote to team channel
# Add content to OneNote 
# Attach a DvT PowerApp to channel with custom parameter
# Add empty Whiteboard to team channel
# Attach existing Whiteboard to team channel
# Attach SharePoint page to team channel
# Create a SharePoint list
# Attach the list to channel tab
# Write into to provisioning list to start a Power Automate
# Create a team in Migration mode and populate some messages there
# Power Automate post provisioning actions
 

##Tenant and App Specific Values
$appconfig = Get-Content -Path "c:\misc\appconfig.txt" | ConvertFrom-Json
$appId = $appconfig.appID
$appSecret = $appconfig.appSecret
$tenantId = $appconfig.tenantId
$OwnerGUID = $appconfig.ownerId
$FinalizeSiteID = $appconfig.FinalizeSiteID
$ProvisionListID = $appconfig.ProvisionListID

$tokenAuthURI = "https://login.microsoftonline.com/$tenantId/oauth2/token"


$requestBody = "grant_type=client_credentials" + 
"&client_id=$appID" +
"&client_secret=" + [URI]::EscapeDataString($appSecret) +
"&resource=https://graph.microsoft.com"

$tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenAuthURI -body $requestBody -ContentType "application/x-www-form-urlencoded"

$accessToken = $tokenResponse.access_token
#Write-Host $accessToken


#Delegated Device Code Token for certain parts requiring delegated permissions
## BIG Thank you Lee Ford for help and this script part! https://www.lee-ford.co.uk/graph-api-device-code/

# Application (client) ID, tenant ID and redirect URI
$clientId = $appId

$resource = "https://graph.microsoft.com/"
$scope = "User.Read.All Group.Read.All ChannelSettings.ReadWrite.All Group.ReadWrite.All" 

$codeBody = @{ 

    resource  = $resource
    client_id = $clientId
    scope     = $scope

}

# Get OAuth Code
$codeRequest = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenantId/oauth2/devicecode" -Body $codeBody

# Print Code to console
Write-Host "`n$($codeRequest.message)"

$tokenBody = @{

    grant_type = "urn:ietf:params:oauth:grant-type:device_code"
    code       = $codeRequest.device_code
    client_id  = $clientId

}

# Get OAuth Token
while ([string]::IsNullOrEmpty($tokenRequest.access_token)) {

    $tokenRequest = try {

        Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenantId/oauth2/token" -Body $tokenBody

    }
    catch {

        $errorMessage = $_.ErrorDetails.Message | ConvertFrom-Json

        # If not waiting for auth, throw error
        if ($errorMessage.error -ne "authorization_pending") {

            throw

        }

    }

}

$delegatedtoken = $tokenRequest.access_token
Write-Host "Done"







## Create Team using a team template
##
## What's the team we are creating?
##

$teamID = ""
$tJSON = ""
$teamname = "GraphTeam " + (get-date).ToString('T').Substring(3, 2)
Write-Host $teamname
$privateChannelName = "Management" 
## Create Team using a custom team template "Demotemplate". Put in your own template.
$templateID = "a664231f-fbd4-4b9f-871d-c11bd066ead2"

##
## Different standard templates: https://docs.microsoft.com/en-us/MicrosoftTeams/get-started-with-teams-templates
## Custome templates are managed in TAC ( Teams Admin Center )
##


$tJSON = '
{
    "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('''+ $templateID + ''')",
    "visibility": "Public",
    "displayName": "' + $teamname + '",
    "description": "Description text that should be filled in..",
    "members":[
      {
         "@odata.type":"#microsoft.graph.aadUserConversationMember",
         "roles":[
            "owner"
         ],
         "user@odata.bind":"https://graph.microsoft.com/v1.0/users('''+ $ownerGUID + ''')"
      }
   ],    
   "memberSettings": {
       "allowCreateUpdateChannels": true,
       "allowDeleteChannels": true,
       "allowAddRemoveApps": true,
       "allowCreateUpdateRemoveTabs": true,
       "allowCreateUpdateRemoveConnectors": true
   },
   "guestSettings": {
       "allowCreateUpdateChannels": false,
       "allowDeleteChannels": false
   },
   "funSettings": {
       "allowGiphy": true,
       "giphyContentRating": "Moderate",
       "allowStickersAndMemes": true,
       "allowCustomMemes": true
   },
   "messagingSettings": {
       "allowUserEditMessages": true,
       "allowUserDeleteMessages": true,
       "allowOwnerDeleteMessages": true,
       "allowTeamMentions": true,
       "allowChannelMentions": true
   },
   "discoverySettings": {
       "showInTeamsSearchAndSuggestions": true
   }
}
' 

##
#Write-Host $tJSON
$graphResponse = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/beta/teams" -Headers @{"Authorization" = "Bearer $accessToken" } -Body $tJSON -ContentType "application/json"
Write-host "Response: " $graphResponse


## Using Team templates with Graph API makes it easier to define channels and applications included. 
## It can take a while to create the team when using Team templates

## Retrieve TeamID by going through all teams.. not ideal but works for the demo 
##
## NOTE! In case there are more than 999 Teams you have two possibilities
##   1. Resolve odata NextLink and do repeated Calls
##   2. Use Get-Team to get the ID (easier)

$groupsListURI = "https://graph.microsoft.com/beta/groups?top=999"
$graphResponse = Invoke-RestMethod -Method Get -Uri $groupsListURI -Headers @{"Authorization" = "Bearer $accessToken" }
Write-Host $graphResponse
foreach ($group in $graphResponse.value) {
  #write-host $group.displayname, $teamname
  if ($group.displayname -eq $teamname ) {
    write-host $group.displayname "=" $group.id 
    $teamID = $group.id
  }
}

## or user Teams PowerShell to find out the teamID! 
## $teamID = ( Get-Team | Where-Object {$_.displayname -match $teamname }).GroupId

write-host $teamID

## Beware of duplicates!

## Get Join link to the Teams to share and view all other settings & info
$TeamWebUrl = ""
$teamURI = "https://graph.microsoft.com/v1.0/teams/" + $teamID 
$graphResponse = Invoke-RestMethod -Method Get -Uri $teamURI -Headers @{"Authorization" = "Bearer $accessToken" }
$TeamWebUrl = $graphResponse.webUrl
Write-Host "Join using this team url" $TeamWebUrl
$PrimaryChannelID = $graphResponse.internalId
## There is also a Graph API call named Get Primary Channel https://docs.microsoft.com/en-us/graph/api/team-get-primarychannel

## view other team settings
$graphResponse

## Update some Team settings. Managing and reading team settings can help to document your environment.
$updateSettingsJSON = '
{  
   "memberSettings": {
     "allowCreateUpdateChannels": false
   },
   "messagingSettings": {
     "allowUserEditMessages": false,
     "allowUserDeleteMessages": false
   },
   "funSettings": {
     "allowGiphy": false
   },
   "discoverySettings": {
     "showInTeamsSearchAndSuggestions": false
   }
 }
'
$graphResponse = Invoke-RestMethod -Method Patch -Uri $teamURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $updateSettingsJSON -ContentType "application/json"
Write-Host $graphResponse

## Check updated settings
$graphResponse = Invoke-RestMethod -Method Get -Uri $teamURI -Headers @{"Authorization" = "Bearer $accessToken" } 
$graphResponse


## Provisiong email to a channel 
## Requires delegated access.. but it is possible
$emailURI = $teamURI + '/channels/' + $PrimaryChannelID + '/provisionEmail'
$graphResponse = Invoke-RestMethod -Method Post  -Uri $emailURI -Headers @{"Authorization" = "Bearer $delegatedToken" }
$emailToGeneral = $graphResponse.email
Write-Host $emailToGeneral
##Write-Host $graphResponse

## https://docs.microsoft.com/en-us/graph/api/channel-provisionemail?view=graph-rest-beta&tabs=http

## all channel settings
$channelURI = $teamURI + '/channels/' + $PrimaryChannelID 
$graphResponse = Invoke-RestMethod -Method Get  -Uri $channelURI -Headers @{"Authorization" = "Bearer $accessToken" }
Write-Host $graphResponse.email

Write-Host $graphResponse

## Update team photo 


## Get photo from existing team
$photoTeamID = ''
$photoURI = 'https://graph.microsoft.com/beta/teams/$photoTeamID/photo/$value'

$graphResponse = Invoke-RestMethod -Method Get -Uri $photoURI -Headers @{"Authorization" = "Bearer $accessToken" }
$photo = $graphResponse

## Put picture to the new team 
## https://docs.microsoft.com/en-us/graph/api/profilephoto-update?view=graph-rest-beta&tabs=http
$photoURI = 'https://graph.microsoft.com/beta/groups/' + $teamId + '/photo/$value'
$graphResponse = Invoke-RestMethod -Method Put -Uri $photoURI -Headers @{"Authorization" = "Bearer $delegatedToken" } -Body $photo -ContentType "image/jpeg"
## Check picture from the SharePoint site, it takes time before picture updates to Teams  


## Add more Applications. 
$appsURI = "https://graph.microsoft.com/beta/teams/" + $teamId + "/installedApps"

## Approvals
$appsJSON = '
{
   "teamsApp@odata.bind": "https://graph.microsoft.com/beta/appCatalogs/teamsApps/7c316234-ded0-4f95-8a83-8453d0876592"
  }'
#Write-Host $appsJSON
$graphResponse = Invoke-RestMethod -Method Post -Uri $appsURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $appsJSON -ContentType "application/json"
Write-Host $graphResponse

$appsURI = "https://graph.microsoft.com/beta/teams/" + $teamId + "/installedApps"


##
## wiki is there.. must make a loop to delete them all
## Get all channels and apply a second command to delete Wiki tab
##
## get channel ids
$newsChannelID = ""
$channelListURI = "https://graph.microsoft.com/v1.0/teams/" + $teamID + "/channels"
$graphResponse = Invoke-RestMethod -Method Get -Uri $channelListURI -Headers @{"Authorization" = "Bearer $accessToken" }

foreach ($channel in $graphResponse.value) {

  $channeltabsURI = "https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $channel.ID + "/tabs?expand=teamsApp"
  $ctabsResponse = Invoke-RestMethod -Method Get -Uri $channelTabsURI -Headers @{"Authorization" = "Bearer $accessToken" }
  foreach ($tab in $ctabsResponse.value) {
    ##Write-Host $channel.displayName $tab.displayName , $tab.teamsApp
    if ($channel.displayName -eq "News 📰📢") {
      $newsChannelID = $channel.id
    }
    ##$tabs.ID
    if ($tab.teamsApp.id -eq "com.microsoft.teamspace.tab.wiki" ) {
      Write-Host $tab.displayName " in " $channel.displayName
      $deleteWiki = "https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $channel.ID + "/tabs/" + $tab.id
      Invoke-RestMethod -Method Delete -Uri $deleteWiki -Headers @{"Authorization" = "Bearer $accessToken" }
      Write-Host "Deleted"       
    }
  }
}


## All wikis removed. 

##
## Copy template Files into Team
## From: Another Team: Graph API Session Templates , General Channel
## Target: General-channel in a new team
##

## Retrieve source and target DriveIDs  
## Get template files from an existing team
$templatefilesourceTeam = "Graph API Session Templates"
## Find source team ID
$groupsListURI = "https://graph.microsoft.com/beta/groups?top=999"
$graphResponse = Invoke-RestMethod -Method Get -Uri $groupsListURI -Headers @{"Authorization" = "Bearer $accessToken" }
Write-Host $graphResponse
foreach ($group in $graphResponse.value) {    
  if ($group.displayname -eq $templatefilesourceTeam ) {
    write-host $group.displayname "=" $group.id 
    $filesourceteamID = $group.id
  }
}

## Find out drive ID (document library) and use general folder in the source team
$filesourceDriveIDURI = "https://graph.microsoft.com/v1.0/groups/" + $filesourceteamID + "/drive"
Write-Host $filesourceDriveIDURI
$graphResponse = Invoke-RestMethod -Method Get -Uri $filesourceDriveIDURI -Headers @{"Authorization" = "Bearer $accessToken" }
##Write-Host $graphResponse
$filesourceDriveID = $graphResponse.id
$sourceFolder = $graphResponse.webUrl + "/general"
Write-Host "Source Folder" $sourceFolder

## get target team drive (document library)
$targetDriveIDURI = "https://graph.microsoft.com/v1.0/groups/" + $teamID + "/drive"
##Write-Host $targetDriveIDURI
$graphResponse = Invoke-RestMethod -Method Get -Uri $targetDriveIDURI -Headers @{"Authorization" = "Bearer $accessToken" }
##Write-Host $graphResponse
$targetDriveID = $graphResponse.id
$targetDocLib = $graphResponse.webUrl

$targetsiteURL = $targetDocLib.Substring(0, $targetDocLib.lastIndexOf('/'))
Write-Host "Target Site URL"  $targetsiteURL
##Write-Host $targetDriveID

##
##
## Make sure each channel folder exist. Create a dummy file to each one.
##
##

## Create Dummy files to all channel folders, this creates the channel

$channelListURI = "https://graph.microsoft.com/v1.0/teams/" + $teamID + "/channels"
$createfileContent = 'Empty file, remove if encountered'
#Write-Host $channelListURI

$graphResponse = Invoke-RestMethod -Method Get -Uri $channelListURI -Headers @{"Authorization" = "Bearer $accessToken" }
foreach ($channel in $graphResponse.value) {
  $createfileURI = "https://graph.microsoft.com/v1.0/drives/" + $targetDriveID + "/root:/" + $channel.displayName + "/dummy_remove.txt:/content"
  Write-Host $createfileURI
  Invoke-RestMethod -Method Put -Uri $createfileURI -Headers @{"Authorization" = "Bearer $accessToken" }  -Body $createfileContent -ContentType "text/plain"
       
}

##
## Cleanup of dummy files would be a proper thing to do.. (implementing that later 😂 ...)
##

## Copy template files from the template team
## Retrieve Target Channel Folder Drive ID 
$targetfoldersURI = "https://graph.microsoft.com/beta/drives/" + $targetDriveID + "/root:/general:/"
##Write-Host $targetfoldersURI
$graphResponse = Invoke-RestMethod -Method Get -Uri $targetfoldersURI -Headers @{"Authorization" = "Bearer $accessToken" }
$targetDriveFolderID = $graphResponse.id
## Copy Template Files - get source uri 
$sourcefilesURI = "https://graph.microsoft.com/beta/drives/" + $filesourceDriveID + "/root:/general:/children"
## Write-Host $sourcefilesURI

$graphResponse = Invoke-RestMethod -Method Get -Uri $sourcefilesURI -Headers @{"Authorization" = "Bearer $accessToken" }
#go through files and copy them to target
foreach ($file in $graphResponse.value) {
  write-host $file.name, $file.id
  $copyURI = "https://graph.microsoft.com/beta/groups/" + $filesourceteamID + "/drive/items/" + $file.id + "/copy"
  $copyJSON = '
        { 
	        "parentReference" : { 
		    "driveId" : "' + $targetDriveID + '", 
		    "id" : "'+ $targetDriveFolderID + '" 
		    }, 
	    "name": "' + $file.name + '" 
        } '
  Write-Host $copyURI
  Write-Host $copyJSON
  $graphCopy = Invoke-RestMethod -Method Post -Uri $copyURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $copyJSON -ContentType "application/json"
  Write-Host $graphCopy
  Write-Host "Copying " + $file.name       
}

## files copied!

## Add a word document as a tab
##
## Get EntityID of Word document  
$wentityID = ""
$wordURI = $targetfoldersURI + "children"
Write-Host $wordURI
$graphResponse = Invoke-RestMethod -Method Get -Uri $wordURI -Headers @{"Authorization" = "Bearer $accessToken" }
# Go through the response and find the right file and get it's EntityID
foreach ($file in $graphResponse.value) {
  Write-Host $file.name
  if ($file.name -eq "Team Practices.docx" ) {
    Write-Host "Found " $file.name 
    Write-Host "ID:" $file.id 
    Write-Host "WebUrl:" $file.webUrl
    $fileweburl = $file.webUrl
    $wentityID = $filewebUrl.Substring($fileweburl.lastIndexOf('sourcedoc=%7B') + 13)
    $wentityID = $wentityID.Substring(0, $wentityid.IndexOf('%7D'))
    Write-Host "Entity ID:" $wentityID
           
  }
}
#write-host $wentityID
$wordJSON = '
{
  "name": "Team Practices",
  "displayName": "Team Practices",
  "teamsAppId": "com.microsoft.teamspace.tab.file.staticviewer.word",
  "configuration": {
     "entityId": "'+ $wentityID + '",
     "contentUrl": "' + $targetDocLib.Replace("%20", " ") + '/general/Team Practices.docx",
     "removeUrl": null,
     "websiteUrl": null
  }
}
'
$pmTabsURI = "https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $PrimaryChannelID + "/tabs"
#Write-Host $pmTabsURI
#Write-Host $wordJSON
$graphResponse = Invoke-RestMethod -Method Post -Uri $pmTabsURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $wordJSON -ContentType "application/json"
Write-Host $graphResponse

##
##
## Add general resources DocLib to General -channel
##
## Add Doclib Tab 

$generalTabsURI = "https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $PrimaryChannelID + "/tabs"
#Write-Host $generalTabsURI

## OLD AppID         "teamsAppId" : "com.microsoft.teamspace.tab.files.sharepoint",        
## OLD application took content & website URL without encoding or changes 
## and you can still use them - but in order to get the full experience on Mobile..
## This is the way! 

$doclibJSON = ' {
  "name": "Resource Documents",
  "displayName": "Resource Documents",
  "teamsAppId": "2a527703-1f6f-4559-a332-d8a7d288cd88",  
  "configuration": {
        "entityId": "",
        "contentUrl": "https://yourtenant.sharepoint.com/sites/ResourceTeams/_layouts/15/teamslogon.aspx?spfx=true&dest=https%3A%2F%2Fnopanen.sharepoint.com%2Fsites%2FResourceTeams%2F_layouts%2F15%2Flistallitems.aspx%3Fapp%3DteamsPage%26listUrl%3D%2Fsites%2FResourceTeams%2FShared%20Documents",
        "removeUrl": "",
        "websiteUrl": "https://yourtenant.sharepoint.com/sites/ResourceTeams/Shared%20Documents" 
    }
  }'

                
#Write-Host $doclibJSON
$graphResponse = Invoke-RestMethod -Method Post -Uri $generalTabsURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $doclibJSON -ContentType "application/json"
Write-Host $graphResponse

##
## Add tenant users or an existing guest users to the team
## If the user is not invited to the tenant - you can use Graph API for that as well https://docs.microsoft.com/en-us/graph/api/invitation-post?view=graph-rest-1.0&tabs=http
## In this demo we add an existing guest user from other tenant

$guestURI = "https://graph.microsoft.com/beta/teams/" + $teamID  + "/members"
$guestGUID = ''
$guestJSON = '
{
  "@odata.type": "#microsoft.graph.aadUserConversationMember",
  "roles": ["member"],
  "user@odata.bind":"https://graph.microsoft.com/v1.0/users('''+ $guestGUID +''')"  
}
'
$graphResponse = Invoke-RestMethod -Method Post -Uri $guestURI -Headers @{"Authorization" = "Bearer $delegatedToken" } -Body $guestJSON -ContentType "application/json"
Write-Host $graphResponse


# Create Tags for the team
## https://docs.microsoft.com/en-us/graph/api/teamworktag-post?view=graph-rest-beta&tabs=http

$tagname = "Graph API Gurus"  ## maximum 40 chars
$tagdesrc = "This TAG was created automatically using Graph API"  ## this does not appear in UI
$tagURI = "https://graph.microsoft.com/beta/teams/" + $teamID  + "/tags"

$tagJSON = '
{
  "displayName": "' + $tagname + '",
  "description": " '+ $tagdesrc +' ",
  "members":[
	{
		"userId":"' + $ownerGUID +'"
	}
  ]
}
'
$graphResponse = Invoke-RestMethod -Method Post -Uri $tagURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $tagJSON -ContentType "application/json"
Write-Host $graphResponse
$tagid = $graphResponse.id

## List Tags
$graphResponse = Invoke-RestMethod -Method Get -Uri $tagURI -Headers @{"Authorization" = "Bearer $accessToken" } 
Write-Host $graphResponse

## Get Tag
$graphResponse = Invoke-RestMethod -Method Get -Uri $tagURI -Headers @{"Authorization" = "Bearer $accessToken" } 
Write-Host $graphResponse.value.displayName $graphResponse.value.description 


# add an other member to a tag. In this demo it is a guest user we just added
$tagaddmemberURI = $tagURI+"/" + $tagid + "/members"
$tagaddJSON ='
{
  "userId":"' + $guestGUID +'"
}'

$graphResponse = Invoke-RestMethod -Method Post -Uri $tagaddmemberURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $tagaddJSON -ContentType "application/json"
Write-Host $graphResponse


## List members
$taglistURI = $tagURI+"/" + $tagid + "/members"
$graphResponse = Invoke-RestMethod -Method Get -Uri $taglistURI -Headers @{"Authorization" = "Bearer $accessToken" } 
Write-Host $graphResponse
foreach ($tagmember in $graphResponse.value) {
  Write-Host $tagmember.userId " " $tagmember.displayName  
}


## Check out more from https://docs.microsoft.com/en-us/graph/api/resources/teamworktag

## Channel moderation settings


## modify News channel to include moderation settings
## General channel can not be updated with moderation settings
$channelJson = '
{   
  "description": "Updated channel moderation in effect",  
  "moderationSettings": {
        "userNewMessageRestriction": "moderators",
        "replyRestriction": "authorAndModerators",
        "allowNewMessageFromBots": true,
        "allowNewMessageFromConnectors": false
    }
}  
'
## userNewMessageRestriction	Indicates who is allowed to post messages to teams channel. Possible values are: everyone, everyoneExceptGuests, moderators, unknownFutureValue.
## replyRestriction	Indicates who is allowed to reply to the teams channel. Possible values are: everyone, authorAndModerators, unknownFutureValue.
## Unfortunately there is no property to manage message Pin in Graph API


$ChannelURI = "https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $newsChannelID
$graphResponse = Invoke-RestMethod -Method Patch -Uri $ChannelURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $channelJson -ContentType "application/json"

## Get channel information
$graphResponse = Invoke-RestMethod -Method Get -Uri $ChannelURI -Headers @{"Authorization" = "Bearer $accessToken" } 
$graphResponse

##
## Create a Private Channel 
Write-Host $privateChannelName
$PrivateChannelID = ""
$channelJson = '
{
    "@odata.type": "#Microsoft.Graph.channel",
    "displayName": "'+ $privateChannelName + '",
    "description": "Channel for '+ $privateChannelName + '",
    "membershipType": "private",
    "members":
     [
        {
           "@odata.type":"#microsoft.graph.aadUserConversationMember",
           "user@odata.bind":"https://graph.microsoft.com/v1.0/users(''' + $ownerGUID + ''')",
           "roles":["owner"]
        }
     ]
}  
'

$ChannelURI = "https://graph.microsoft.com/beta/teams/" + $teamID + "/channels"
#Write-Host $ChannelURI
#Write-Host $ChannelJson
$graphResponse = Invoke-RestMethod -Method Post -Uri $ChannelURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $channelJson -ContentType "application/json"
$PrivateChannelID = $graphResponse.id
Write-host $PrivateChannelID



##
## Add OneNote 
##
## Create OneNote Notebook
## Note: Notes.ReadWrite.All required

$notebookURI = "https://graph.microsoft.com/v1.0/groups/" + $teamId + "/onenote/notebooks"
$notebookcreateJSON = '
 {
	"displayName": "'+ $teamname + ' Notebook"
   }
'
Write-Host $notebookURI
Write-Host $notebookcreateJSON
$graphResponse = Invoke-RestMethod -Method Post -Uri $notebookURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $notebookcreateJSON -ContentType "application/json"
Write-Host $graphResponse

$notebookID = $graphResponse.id
$sectionsURL = $graphResponse.sectionsUrl
Write-Host $notebookID
Write-Host $sectionsURL


##
## Add OneNote Tab 
##
$onenotewebUrl = ""
## Retrieve onenote web url 
$graphResponse = Invoke-RestMethod -Method Get -Uri $notebookURI -Headers @{"Authorization" = "Bearer $accessToken" }
#Write-Host $graphResponse
foreach ($notebook in $graphResponse.value) {
  #Write-Host $notebook
  if ($notebook.id -eq $notebookID ) {
    $notebook.displayName
    Write-Host $notebook.links.oneNoteWebUrl.href
    $onenotewebUrl = $notebook.links.oneNoteWebUrl.href
  }
}

#Write-Host $onenotewebURL
$notebookeName = $onenotewebUrl.Substring($onenotewebUrl.lastIndexOf('/') + 1)
$onenotewebUrl = $onenotewebUrl.Replace("/", "%2F")
$onenotewebUrl = $onenotewebUrl.Replace(":", "%3A")
$sectionsURL = $sectionsURL.Replace("/", "%2F")
$sectionsURL = $sectionsURL.Replace(":", "%3A")
#Random GUID & build URL
$guid = [System.Guid]::NewGuid()
$entityID = $guid.Guid + "_" + $notebookID
$websiteURL = "https://www.onenote.com/teams/TabRedirect?redirectUrl=" + $onenotewebUrl 
#Write-Host $onenotewebUrl
$contentURL = "https://www.onenote.com/teams/TabContent?notebookSource=Pick&notebookSelfUrl=https%3A%2F%2Fwww.onenote.com%2Fapi%2Fv1.0%2FmyOrganization%2Fgroups%2F" + $teamID + "%2Fnotes%2Fnotebooks%2F" + $notebookID + "&oneNoteWebUrl=" + $onenotewebUrl + "&notebookName=" + $notebookeName + "&siteUrl=" + $targetsiteURL + "&createdTeamType=Standard&ui={locale}&tenantId={tid}&upn={userPrincipalName}&groupId={groupId}&theme={theme}"
$removeURL = "https://www.onenote.com/teams/TabRemove?notebookSource=Pick&notebookSelfUrl=https%3A%2F%2Fwww.onenote.com%2Fapi%2Fv1.0%2FmyOrganization%2Fgroups%2F" + $teamID + "%2Fnotes%2Fnotebooks%2F" + $notebookID + "&oneNoteWebUrl=" + $onenotewebUrl + "&notebookName=" + $notebookeName + "&siteUrl=" + $targetsiteURL + "&createdTeamType=Standard&ui={locale}&tenantId={tid}&upn={userPrincipalName}&groupId={groupId}&theme={theme}"
$notebookJSON = '
 {
	"name": "'+ $notebookeName + '",
	"displayName": "'+ $notebookeName.Replace("%20", " ") + '",
    "teamsAppId": "0d820ecd-def2-4297-adad-78056cde7c78",
        "configuration": {
        	"entityId": "'+ $entityID + '",
		    "contentUrl": "'+ $contentURL + '",
		    "removeUrl": "'+ $removeURL + '",
		    "websiteUrl": "'+ $websiteURL + '"      
      }
 }
'
#Write-Host $generalTabsURI
#Write-Host $notebookJSON
$notebookURI 
$graphResponse = Invoke-RestMethod -Method Post -Uri $generalTabsURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $notebookJSON -ContentType "application/json"
Write-Host $graphResponse

##
## can we write something to that notebook?
##
## Add Section
$addSectionURL = $notebookURI + "/" + $notebookID + "/sections"
$addSectionJSON = '
{
    "displayName": "ECS Graph API Fun"
 }'
$graphResponse = Invoke-RestMethod -Method Post -Uri $addSectionURL -Headers @{"Authorization" = "Bearer $accessToken" } -Body $addSectionJSON -ContentType "application/json"
$newSectionID = $graphResponse.id
## Add Page
$addPageURL = "https://graph.microsoft.com/beta/groups/" + $teamId + "/onenote/sections/" + $newSectionID + "/pages"
$pageHtml = '
<html>
  <head>
    <title>CollabSummit <b>Graph API </n> in ' + $teamname + '</title>
  </head>
  <body>
    <p>Fun = Experimenting with Graph API and PowerShell in Teams!</p>
  </body>
</html>
'
$graphResponse = Invoke-RestMethod -Method Post -Uri $addPageURL -Headers @{"Authorization" = "Bearer $accessToken" } -Body $pageHtml -ContentType "text/html"
Write-Host $graphResponse

## 
## Add a predefined PowerApps application tab with some custom parameter
## Make sure Dataverse for Teams PowerApp is distributed as application to the org first 
## AND don't forget to give permisssions either (in Dataverse for Teams & Power Apps)

## Best way to retrieve JSON required is to use Graph Explorer and retrieve channel Tabs
## https://docs.microsoft.com/en-us/graph/api/channel-list-tabs

## Note: teamsApp@odata.bind ID is different than in the manifest ID. 
## Retrieve applications installed to Teams 
## match the one we need or retrieve the value from TAC Applications

## https://docs.microsoft.com/en-us/graph/api/appcatalogs-list-teamsapps
## Application permissions are not supported.
## Use Graph Explorer
## GET https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=distributionMethod eq 'organization'
## AppCatalog.Submit, AppCatalog.Read.All, AppCatalog.ReadWrite.All
##    "value": [
##  {
##    "id": "f612ad24-3333-44ee-a8a5-f85e627c50a6",
##    "externalId": "35b6a3b1-7676-4141-ac75-bd580cc408a1",
##    "displayName": "TeamsDemo",
##    "distributionMethod": "organization"
##},

## ID : Application ID in Teams Instance 
## EnternalID: Application ID in the manifest, "PowerApps App ID"

## Add application to the team
$appsJSON = '
{
   "teamsApp@odata.bind": "https://graph.microsoft.com/beta/appCatalogs/teamsApps/f612ad24-3333-44ee-a8a5-f85e627c50a6"
  }'
#Write-Host $appsJSON
$graphResponse = Invoke-RestMethod -Method Post -Uri $appsURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $appsJSON -ContentType "application/json"
Write-Host $graphResponse



## Use subEntityID for delivering a custom parameter to the app in that team
$subEntityID = "Enterprise 1701"

$powerTabJSON = '  {
            "id": "GraphDemoTab1",
            "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/f612ad24-3333-44ee-a8a5-f85e627c50a6",            
            "displayName": "Graph Power Demo",            
            "webUrl": "webUrl to app",
            "configuration": {
                "entityId": "35b6a3b1-7618-41f1-ac75-ce580cc408a1",
                "contentUrl": "Content URL to app but remember &subEntityId=' + $subEntityID +'&teamId={teamId}&teamType={teamType}&theme={theme}&userTeamRole={userTeamRole}",
                "removeUrl": null,
                "websiteUrl": "websiteurl"
            }
} '
$powerTabURI = "https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $PrimaryChannelID + "/tabs"
##Write-Host $powerTabURI
##Write-Host $powerTabJSON
$graphResponse = Invoke-RestMethod -Method Post -Uri $powerTabURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $powerTabJSON -ContentType "application/json"
Write-Host $graphResponse


##
## Add Whiteboard to a team channel as a tab to General
## Template added Whiteboard can be found at Chit Chat channel. 


$whiteboardJSON = '
{
  ”displayName”: ”Whiteboard via Graph API”,
  ”teamsApp@odata.bind”: ”https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/95de633a-083e-42f5-b444-a4295d8e9314”,
  ”configuration”: {
               ”entityId”: null,
               ”contentUrl”: null,
               ”removeUrl”: null,
               ”websiteUrl”: null
  }
}
'
$PrimaryTabsURI = "https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $PrimaryChannelID + "/tabs"
$graphResponse = Invoke-RestMethod -Method Post -Uri $PrimaryTabsURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $whiteboardJSON -ContentType "application/json"
Write-Host $graphResponse

## Of course the Whiteboard needs to be set up separately in this way
## There is a way to add existing Whiteboard to the team via Graph API

## Let's add an existing Whiteboard. 
$whiteboardJSON = '
{
  ”displayName”: ”ECS Whiteboard”,
  ”teamsApp@odata.bind”: ”https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/95de633a-083e-42f5-b444-a4295d8e9314”,
  ”configuration”: {
    "entityId": "entityid to existing wb",
    "contentUrl": "content url to existing wb",
    "removeUrl": null,
    "websiteUrl": "website url to existing wb"
  }
}
'

$PrimaryTabsURI = "https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $PrimaryChannelID + "/tabs"
$graphResponse = Invoke-RestMethod -Method Post -Uri $PrimaryTabsURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $whiteboardJSON -ContentType "application/json"
Write-Host $graphResponse

## Remember to add permissions to all team members by a sharing link via OneDrive


## Add SharePoint Page as tab
##
## Hmm, SharePoint page configuration is not supported according to Docs.Microsoft.com.  
## https://docs.microsoft.com/en-us/graph/teams-configuring-builtin-tabs
## *SharePoint page and list tabs* "Configuration is not supported. If you want to configure the tab, consider using a Website tab."
##
## What about if we add the home page anyway as a tab?
##

$spcontentURL = ""
$spWebsiteUrl = ""
$spcontentURL = $targetsiteURL + '/_layouts/15/teamslogon.aspx?spfx=true&dest='
## Content URL must be formatted differently
## "contentUrl": "https://{siteaddress}/_layouts/15/teamslogon.aspx?spfx=true&dest={encodedsiteaddress}%2FSitePages%2FHome.aspx",
$spencodedURL = $targetsiteURL
$spencodedURL = $spencodedURL.Replace("/", "%2F")
$spencodedURL = $spencodedURL.Replace(":", "%3A")
$spcontentURL = $spcontentURL + $spencodedURL
$spcontentURL = $spcontentURL + '%2FSitePages%2FHome.aspx'
$spWebsiteUrl = $targetsiteURL + '/SitePages/Home.aspx'
## we could copy a suitable page there and add it as a tab by just knowing it's name
## in this case we use Home.aspx because it is there on default
Write-Host $spcontentURL
Write-Host $spWebsiteUrl

$spJSON = '
{
"teamsAppId": "2a527703-1f6f-4559-a332-d8a7d288cd88",
"name": "SharePoint tab",
"displayName": "SharePoint Home page",	
"sortOrderIndex": "10000",	
"configuration": {
        "contentUrl": "'+ $spcontentURL + '",
        "websiteUrl": "' + $spWebsiteUrl + '",          	
}}
'

$piTabsURI = "https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $PrimaryChannelID + "/tabs"
$graphResponse = Invoke-RestMethod -Method Post -Uri $piTabsURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $spJSON -ContentType "application/json"
Write-Host $graphResponse

## 
## Create a list in SharePoint 
## Unfortunately we can't use Microsoft List templates for this one. So we do it the hard way...

## We need SiteID first. We could have been smart and used this when we figured out the target site url (weburl returned by this call). 
$siteURI = "https://graph.microsoft.com/beta/groups/" + $teamID + "/sites/root"
Write-Host $siteURI
$graphResponse = Invoke-RestMethod -Method Get -Uri $siteURI -Headers @{"Authorization" = "Bearer $accessToken" }
#Write-Host $graphResponse
$siteID = $graphResponse.id
Write-Host $siteID

$listName = "ECS List"
$createListURI = "https://graph.microsoft.com/beta/sites/" + $siteID + "/lists"

$createListJSON = '
{ 
	"displayName": "'+ $listName + '", 
	"columns": [ 
		{ 
			"name": "Entry date", 
			"dateTime": {
			   "displayAs": "standard",
			   "format": "dateTime"
			  }			
		},
		 {   
            "name": "DaysNumbered", 
			"displayName": "Length in days",
			"number": {
			    "decimalPlaces": "none",
			    "displayAs": "number",
			    "maximum": 1.7976931348623157e+308,
			    "minimum": -1.7976931348623157e+308
			   }
		},
		  {
		    "name": "Information",
		    "text": {
		        "allowMultipleLines": false,
		        "appendChangesToExistingText": false,
		        "linesForEditing": 0,
		        "maxLength": 255
		       }
		},
		  {
		    "name": "Source",
		    "text": {
		        "allowMultipleLines": false,
		        "appendChangesToExistingText": false,
		        "linesForEditing": 0,
		        "maxLength": 255
		       }
		  }
	 ], 
	"list": 
	{ "template": "genericList" } 
}
'
#Write-Host $createListJSON
#Write-Host $createListURI
$graphResponse = Invoke-RestMethod -Method Post -Uri $createListURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $createListJSON -ContentType "application/json"
#write-host $graphResponse
$listID = $graphResponse.id
$listWebUrl = $graphResponse.webUrl
Write-Host  "ListID: " $listID
 
## Add list as tab
$spcontentURL = "" 
$spencodedURL = $targetsiteURL
$spencodedURL = $spencodedURL.Replace("/", "%2F")
$spencodedURL = $spencodedURL.Replace(":", "%3A")
$encodedListName = $listName.Replace(" ", "%20")
$spcontentURL = $targetsiteURL + "/_layouts/15/teamslogon.aspx?spfx=true&dest=" + $spencodedURL + "%2FLists%2F" + $encodedlistName + "%2FAllItems.aspx%3Fp%3D11"
$splistwebURL = $listWebUrl + '/AllItems.aspx?p=11'
    
$spJSON = '
    {
        "teamsAppId": "2a527703-1f6f-4559-a332-d8a7d288cd88",
        "name": "'+ $listName + '",
        "displayName": "'+ $listName + '",	
        "sortOrderIndex": "10000",	
        "configuration": {
	            "contentUrl": "'+ $spcontentURL + '",
                "websiteUrl": "' + $splistwebUrl + '",          	
    }}
    '


$graphResponse = Invoke-RestMethod -Method Post -Uri $generalTabsURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $spJSON -ContentType "application/json"
Write-Host $graphResponse


## Let's finalize the provisioning by writing the new team name and ID to a list for Power Automate to get started
## The post-provisioning Power Automate posts a welcome message and
## adds a Planner, with couple of tasks, to the team and adds it as a tab


$finalizeURL= "https://graph.microsoft.com/v1.0/sites/"+ $FinalizeSiteID+ "/lists/"+ $ProvisionListID+"/items"
$finalizeJSON = '
{
  "fields": {
    "Title": "'+ $teamname +'",
    "Status": "Provisioning",
    "TeamID": "'+ $teamID +'"
  }
}
'
$graphResponse = Invoke-RestMethod -Method Post -Uri $finalizeURL -Headers @{"Authorization" = "Bearer $accessToken" } -Body $finalizeJSON -ContentType "application/json"
Write-Host $graphResponse




## Migration 
## Create a team in migration mode
## Allows to import / create messages if needed
$migrateTeamName = 'Migrated team ' + (get-date).ToString('T').Substring(3, 2)
$migrateJSON = '
{
    "@microsoft.graph.teamCreationMode": "migration",
    "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates(''standard'')",
    "displayName": "' + $migrateTeamName + '",
    "description": "Sample Team Description",
    "createdDateTime": "2021-10-10T11:22:17.067Z",
    "channels": [
    {
      "displayName": "General",
      "description": "Migrated team general channel",
      "@microsoft.graph.channelCreationMode": "migration",
      "createdDateTime": "2020-10-10T11:11:11.047Z"  
    }

    ]
}
'

#Write-Host $migrateJSON
$graphResponse = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/teams" -Headers @{"Authorization" = "Bearer $accessToken" } -Body $migrateJSON -ContentType "application/json"
Write-host "Response: " $graphResponse


## Talk for a moment - create a brief  waiting/delay before progressing

$groupsListURI = "https://graph.microsoft.com/beta/groups?top=999"
$graphResponse = Invoke-RestMethod -Method Get -Uri $groupsListURI -Headers @{"Authorization" = "Bearer $accessToken" }
Write-Host $graphResponse
foreach ($group in $graphResponse.value) {
    
  if ($group.displayname -eq $migrateTeamName ) {
    write-host $group.displayname = $group.id 
    $migrateteamID = $group.id
  }
}

write-host $migrateteamID

## Get General channel and create a other channel
$teamURI = "https://graph.microsoft.com/v1.0/teams/" + $migrateteamID 
$graphResponse = Invoke-RestMethod -Method Get -Uri $teamURI -Headers @{"Authorization" = "Bearer $accessToken" }
$graphResponse 
$MigratedPrimaryChannelID = $graphResponse.internalId

$migratedChannelJSON = '
{
  "@microsoft.graph.channelCreationMode": "migration",
  "displayName": "Community Champions",
  "description": "This channel is for community champions program discussion",
  "membershipType": "standard",
  "createdDateTime": "2020-10-10T11:25:17.047Z"  
}
'

$ChannelURI = $teamURI + "/channels"
$graphResponse = Invoke-RestMethod -Method Post -Uri $ChannelURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $migratedChannelJSON -ContentType "application/json"
Write-host "Response: " $graphResponse
$migratedChannelID = $graphResponse.id

## Write Messages to Migrated Champions Channel
$MigratedChannelURI = $teamURI + "/channels/" + $MigratedChannelID + "/messages"

$messageJSON = '
{ 
	"createdDateTime":"2021-10-10T18:00:00.000Z",
    "from": {
        "user": {
            "id": "' + $ownerGUID + '",
            "displayName":"Vesku Admin",
            "userIdentityType":"aadUser"
        }
    },
    "importance": "high",
	"subject": "Welcome to the ECS Demo Team",
	"body": 
		{ 
			"contentType": "html", 
			"content": "Hello team! Welcome! Read me first, like all instructions!"
		} 	
}
'
#Write-Host $messageJSON

$graphResponse = Invoke-RestMethod -Method Post -Uri $MigratedChannelURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $messageJSON -ContentType "application/json"
Write-host "Response: " $graphResponse

## write  other message to Champions channel
$messageJSON = '
{ 
	"createdDateTime":"2020-11-01T18:00:00.000Z",
    "from": {
        "user": {
            "displayName":"CEO Admin"
        }
    },
    "importance": "high",
	"subject": "Welcome to Champions. Read me first!",
	"body": 
		{ 
			"contentType": "html", 
			"content": "Hello team! Welcome!"
		} 	
}
'

$graphResponse = Invoke-RestMethod -Method Post -Uri $MigratedChannelURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $messageJSON -ContentType "application/json"
Write-host "Response: " $graphResponse


## Let's add a message with inline picture!
$imagefile = [convert]::ToBase64String((Get-Content "C:\misc\teampic.png" -Encoding Byte))

$message2JSON = '
{
    "createdDateTime":"2020-11-11T18:00:00.000Z",
    "from": {
        "user": {
            "displayName":"Marketing VP"
        }
    },
    "subject": "ECS 2021",
    "body": {
          "contentType": "html",
          "content": "<div><B>Hello everyone and Enjoy Graph API at ECS 2021!<div>\n<div><span><img height=\"250\" src=\"../hostedContents/1/$value\" width=\"176\" style=\"vertical-align:bottom; width:176px; height:250px\"></span>\n\n</div>\n\n\n</div>Rock and Roll!\n</div>"
      },
      "hostedContents":[
          {
              "@microsoft.graph.temporaryId": "1",
              "contentBytes": "'+ $imagefile+'",
              "contentType": "image/png"
          }
      ]
  }
  '

$graphResponse = Invoke-RestMethod -Method Post -Uri $MigratedChannelURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $message2JSON -ContentType "application/json"
Write-host "Response: " $graphResponse

## End Migration for Champions channel
$EndMigrationURI = $teamURI + "/channels/" + $migratedChannelID + "/completeMigration"
$graphResponse = Invoke-RestMethod -Method Post -Uri $EndMigrationURI -Headers @{"Authorization" = "Bearer $accessToken" } 
#Write-host "Response: " $graphResponse

## End Migration for General Channel
$EndMigrationURI = $teamURI + "/channels/" + $migratedPrimaryChannelID + "/completeMigration"
$graphResponse = Invoke-RestMethod -Method Post -Uri $EndMigrationURI -Headers @{"Authorization" = "Bearer $accessToken" } 
#Write-host "Response: " $graphResponse

## End Migration for team
$EndTeamMigrationURI = $teamURI + "/completeMigration"

$graphResponse = Invoke-RestMethod -Method Post -Uri $EndTeamMigrationURI -Headers @{"Authorization" = "Bearer $accessToken" } 
#Write-host "Response: " $graphResponse

## Add members and owners
$addOwnerJSON = '
{
    "@odata.type":"#microsoft.graph.aadUserConversationMember",
    "roles":[
       "owner"
    ],
    "user@odata.bind":"https://graph.microsoft.com/v1.0/users('''+ $ownerGUID + ''')"
 }
 '
$ownerURI = $teamURI + "/members"
$graphResponse = Invoke-RestMethod -Method Post -Uri $ownerURI -Headers @{"Authorization" = "Bearer $accessToken" } -Body $addOwnerJSON -ContentType "application/json"
Write-host "Response: " $graphResponse 


## https://docs.microsoft.com/en-us/microsoftteams/platform/graph-api/import-messages/import-external-messages-to-teams

## Let's check how our migrated team looks like now..

## Lets check out the Flow demo -part!
## Write to a list to activate a Flow 
## or could we excute it via Graph API.... 


## EOF ##