## Vesa Nopanen, MVP 
## @vesanopanen
## Check out my blog https://myteamsday.com 

## Listing teams & channels and OneNote creation & modifications can be done using Application permissions
## Reading channel messages required delegated permissions

## Using PNPOnline for delegated access. Install PNPOnline first. 
## This one gives a lots of permissions - not all are needed in this OneNote demo. Adjust as needed.
Connect-PnPOnline -Url https://{yourtenantnamehere}.sharepoint.com -Scopes "User.Read","User.ReadBasic.All", "Group.Read.All", "Group.ReadWrite.All", "Files.Read.All", "Files.ReadWrite.All", "Sites.Read.All", "Sites.ReadWrite.All", "Notes.Read.All", "Notes.Create",  "Notes.ReadWrite.All", "Calendars.ReadWrite", "Chat.Read", "Chat.ReadWrite",  "Directory.Read.All", "Directory.ReadWrite.All", "Directory.AccessAsUser.All"
$delegatedaccessToken =Get-PnPAccessToken

## to make things easier, we can lookup team and channel by names.
$teamName ="NCC-1701 crew"
$channelname ="Backup these"

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

## let's create a Notebook to team
$notebookURI = "https://graph.microsoft.com/beta/groups/" + $teamId + "/onenote/notebooks"
$notebookcreateJSON = '
 {
	"displayName": "Export Notebook"
   }
'
Write-Host $notebookURI
Write-Host $notebookcreateJSON
$graphResponse = Invoke-RestMethod -Method Post -Uri $notebookURI -Headers @{"Authorization"="Bearer $delegatedaccessToken"} -Body $notebookcreateJSON -ContentType "application/json"
Write-Host $graphResponse
$notebookID = $graphResponse.id

## adding a section
$SectionURL = $notebookURI + "/" + $notebookID + "/sections"
$SectionJSON ='
{
    "displayName": "Backup of Channel Demo"
 }'
$graphResponse = Invoke-RestMethod -Method Post -Uri $SectionURL -Headers @{"Authorization"="Bearer $delegatedaccessToken"} -Body $SectionJSON -ContentType "application/json"
$SectionID = $graphResponse.id

## adding a new page
$addPageURL = "https://graph.microsoft.com/beta/groups/" + $teamId + "/onenote/sections/" + $SectionID +"/pages"
$pageHtml = '
<html>
  <head>
    <title>Messages in '+ $teamName +'</title>
  </head>
  <body>
    <p><b>This is what has been said in the <b>'+ $channelname +'</b> channel</b></p>'

$messagesURI = "https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $ChannelID + "/messages"
$graphResponse = Invoke-RestMethod -Method Get -Uri $messagesURI  -Headers @{"Authorization"="Bearer $delegatedaccessToken"}
#go through files and copy them to target

foreach ($message in $graphResponse.value)
{
        $messageID = $message.id 
        $pageHtml = $pageHtml + '<p>' + $message.createdDateTime+" "+ $message.from.user.displayName +" <b>" + $message.subject+"</b>:"+$message.body.content +'</p>'
        $repliesURI = "https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $ChannelID + "/messages/" + $messageID + "/replies"

        $repliesResponse = Invoke-RestMethod -Method Get -Uri $repliesURI  -Headers @{"Authorization"="Bearer $delegatedaccessToken"}
        foreach ($reply in $repliesResponse.value ) {
             $pageHtml = $pageHtml + '<p>&nbsp; &nbsp; &nbsp; &nbsp; ' + $reply.createdDateTime+" reply: "+ $reply.from.user.displayName +" " + $reply.subject+":"+$reply.body.content +'</p>'

        }
     $pageHtml = $pageHtml + '<p>---------------------------------------------------------------------------------------------</p>'
       
}

$pageHtml = $pageHtml + '  </body>
</html>
'

$graphResponse = Invoke-RestMethod -Method Post -Uri $addPageURL -Headers @{"Authorization"="Bearer $delegatedaccessToken"} -Body $pageHtml -ContentType "text/html"

