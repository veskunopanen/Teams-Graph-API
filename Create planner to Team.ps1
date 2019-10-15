## Vesa Nopanen, MVP 
## @vesanopanen
## Check out my blog https://myteamsday.com 

## Create a Planner and add that to a team can be done using Application permissions
## Adding also a single bucket and task into that

## Using PNPOnline for delegated access. Install PNPOnline first. 
## This one gives a lots of permissions - not all are needed in this OneNote demo. Adjust as needed.
$tenantname = 'mytenantname'
$pnpURL = 'https://' + $tenantname+'.sharepoint.com'
Connect-PnPOnline -Url $pnpURL -Scopes "User.Read","User.ReadBasic.All", "Group.Read.All", "Group.ReadWrite.All", "Files.Read.All", "Files.ReadWrite.All", "Sites.Read.All", "Sites.ReadWrite.All", "Notes.Read.All", "Notes.Create",  "Notes.ReadWrite.All", "Calendars.ReadWrite", "Chat.Read", "Chat.ReadWrite",  "Directory.Read.All", "Directory.ReadWrite.All", "Directory.AccessAsUser.All"
$delegatedaccessToken =Get-PnPAccessToken

## to make things easier, we can lookup team and channel by names.
$teamName ="NCC-1701 crew"
$channelname ="general"

$groupsListURI = "https://graph.microsoft.com/v1.0/me/joinedTeams"
$graphResponse = Invoke-RestMethod -Method Get -Uri $groupsListURI -Headers @{"Authorization"="Bearer $delegatedaccessToken"}
foreach ($group in $graphResponse.value)
{
    #write-host $group.displayname, $teamName
    if ($group.displayname -eq $teamName ) {
        write-host $group.displayname = $group.id 
        $teamID = $group.id
        }
}

$ChannelID = ""
$channelListURI = "https://graph.microsoft.com/v1.0/teams/" + $teamID + "/channels"
$graphResponse = Invoke-RestMethod -Method Get -Uri $channelListURI -Headers @{"Authorization"="Bearer $delegatedaccessToken"}

foreach ($channel in $graphResponse.value)
    {
            if ($channel.displayName -eq $channelname) {
                $ChannelID = $channel.id
            }
}

write-host "Channel ID:" $ChannelID

## Create a Planner to team
$createPlanUri ="https://graph.microsoft.com/beta/planner/plans"
$createPlanJSON ='{ 
	"owner": "' +$teamID +'", 
	"title": "' +$teamName +' Planner" 
}
' 
$graphResponse = Invoke-RestMethod -Method Post -Uri $createPlanUri -Headers @{"Authorization"="Bearer $delegatedaccessToken"} -Body $createPlanJSON -ContentType "application/json"
$planID = $graphResponse.id
Write-Host "Planner Plan ID: " $planID

##
## add some planner buckets & tasks
##
$bucketJSON ='
{
  "name": "Backlog bucket",
  "planId": "' +$planID + '",
  "orderHint": " !"
}
'
$bucketURI = "https://graph.microsoft.com/beta/planner/buckets"
$graphResponse = Invoke-RestMethod -Method Post -Uri $BucketUri -Headers @{"Authorization"="Bearer $delegatedaccessToken"} -Body $BucketJSON -ContentType "application/json"

$backlockBucketID = $graphResponse.id
Write-Host "Bucket ID: " $backlockBucketID

## Add a Task
$taskJSON ='
{
  "title": "Sample task",
  "planId": "' +$planID + '",
  "bucketId" : "' + $backlockBucketID + '",
  "orderHint": " !",
  "priority" : 2,
  "appliedCategories": {
    "category1": true
  }
}'

$taskURI = "https://graph.microsoft.com/beta/planner/tasks"
$graphResponse = Invoke-RestMethod -Method Post -Uri $TaskUri -Headers @{"Authorization"="Bearer $delegatedaccessToken"} -Body $TaskJSON -ContentType "application/json"
$taskID = $graphResponse.id
Write-Host "TaskID: " $taskID

## Add Planner Tab to a team
##
## https://docs.microsoft.com/en-us/graph/teams-configuring-builtin-tabs

$generalTabsURI = "https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $ChannelID + "/tabs"

$planURL = 'https://tasks.office.com/'+ $tenantname +'.onmicrosoft.com/Home/PlannerFrame?page=7&planId='+ $planID

$plannerJSON =' {
	"name": "'+ $teamname +' Backlog",
	"displayName": "'+  $teamname  +' Backlog",
     "teamsAppId" : "com.microsoft.teamspace.tab.planner",
     "configuration": {
        "entityId": "'+ $planID +'",
		"contentUrl": "' + $planURL+ '",
		"removeUrl": "' + $planURL+ '",
		"websiteUrl": "' + $planURL+ '"
           }
        }'
        
#Write-Host $plannerJSON
$graphResponse = Invoke-RestMethod -Method Post -Uri $generalTabsURI -Headers @{"Authorization"="Bearer $delegatedaccessToken"} -Body $plannerJSON -ContentType "application/json"
Write-Host $graphResponse
