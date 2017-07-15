<###########################################
## Karloch's PowerShell Trello New Sprint ##
############################################>

param ([Parameter(Mandatory=$true)]
    [string]$AppKey,
    [Parameter(Mandatory=$true)]
    [string]$Token,
    [Parameter(Mandatory=$true)]
    [string]$SprintBoardId,
    [Parameter(Mandatory=$true)]
    [string]$NewBoardTitle,
    [string]$OrgId="myorgid",
    [switch]$SendMail=$false,
    [string]$MailFrom="Trello New Sprinter Automation",
    [string[]]$MailTo=@("destinationemail@domain.com"),
    [string]$MailSubject = "New kanban created. Welcome to sprint $(Get-Date -Format "yyyy.MM")!",
    [string]$SmtpServer="smtp.office365.com",
    [int]$SmtpPort=587,
    [switch]$SmtpSsl=$true,
    [System.Management.Automation.PSCredential]$SmtpCred
)
Write-Host -ForegroundColor Yellow "Trello PowerShell New Sprinter v1.0 by Carlos Milán Figueredo - https://calnus.com"
Write-Host ""

# First, let's connect to the Trello REST API and retrieve the board cards and stuff
# We will show progress meanwhile in case user has slow connection
Write-Progress -Activity "Creating new Trello sprint..." -Status "Getting board information..." -PercentComplete 0 
$SprintBoard = Invoke-RestMethod "https://api.trello.com/1/board/$SprintBoardId`?key=$AppKey&token=$Token" -ErrorAction Stop
Write-Progress -Activity "Creating new Trello sprint..." -Status "Getting cards..." -PercentComplete 5
$SprintBoardCardsRaw=Invoke-RestMethod "https://api.trello.com/1/board/$SprintBoardId/cards?key=$AppKey&token=$Token" -ErrorAction Stop
Write-Progress -Activity "Creating new Trello sprint..." -Status "Getting lists..." -PercentComplete 25
$SprintBoardLists=Invoke-RestMethod "https://api.trello.com/1/board/$SprintBoardId/lists?key=$AppKey&token=$Token" -ErrorAction Stop
Write-Progress -Activity "Creating new Trello sprint..." -Status "Getting members..." -PercentComplete 30
$SprintBoardMembers=Invoke-RestMethod "https://api.trello.com/1/board/$SprintBoardId/members?key=$AppKey&token=$Token" -ErrorAction Stop

# In order to ease the process, we are now making a copy of the current board in
# a new one for next sprint. Everything except membership is carried over.
$CreateNewBoardJson = ConvertTo-Json (@{'name'=$NewBoardTitle;'idOrganization'=$OrgId;'prefs_permissionLevel'='org';
    'idBoardSource'=$SprintBoard.id;'keepFromSource'='all'})
Write-Progress -Activity "Creating new Trello sprint..." -Status "Copying board $($SprintBoard.name) to $NewBoardTitle..." -PercentComplete 35
$newBoard=Invoke-RestMethod -Method Post -Uri "https://api.trello.com/1/boards?key=$AppKey&token=$Token" -ContentType "application/json" -Body $CreateNewBoardJson -ErrorAction Stop

# Now, we create the "To Next Srint" list in the old Sprint kanban
$ToNextSprintListJson = ConvertTo-Json(@{'name'='To Next Sprint';'idBoard'=$SprintBoard.id;'pos'='bottom'})
Write-Progress -Activity "Creating new Trello sprint..." -Status "Creating To Next Sprint..." -PercentComplete 70
$ToNextSprintResult=Invoke-RestMethod -Method Post -Uri "https://api.trello.com/1/lists?key=$AppKey&token=$Token" -ContentType "application/json" -Body $ToNextSprintListJson -ErrorAction Stop

# Membership was not carriend over during board copy, we are adding it now
# The board owner is already a member, so I won't add myself
Write-Progress -Activity "Creating new Trello sprint..." -Status "Adding members..." -PercentComplete 75
foreach ($member in $SprintBoardMembers) {
    if($member.username -ne "carlosmilanfigueredo") {
        $memberJson=ConvertTo-Json (@{'idMember'=$member.id;'type'='normal'})
        Invoke-RestMethod -Method Put -Uri "https://api.trello.com/1/boards/$($newBoard.id)/members/$($member.id)`?key=$AppKey&token=$Token" -ContentType "application/json" -Body $memberJson
    }
}

# We define a Hashtable of Arrays :) The idea is to have something like
# SprintBoardCards['In progress'][0] and get a card in the 'In Progress'
# column
$SprintBoardCards = @{}

foreach($list in $SprintBoardLists)
{
    $SprintBoardCards.Add($list.name, @())
}

# Let's get the 2 dimensional array populated. Sadly, Trello API doesn't offer
# lists names in the /1/board/$SprintBoardId/cards, so we will have to play a bit
# with the /1/board/$SprintBoardId/lists
foreach($card in $SprintBoardCardsRaw)
{
    foreach($list in $SprintBoardLists)
    {
        if($card.idList -eq $list.id) {
            $SprintBoardCards[$list.name] += $card
        }
    }
}

# We delete the cards in the Inbox list. They are already in the new board.
Write-Progress -Activity "Creating new Trello sprint..." -Status "Deleting Inbox cards from $($SprintBoard.name)" -PercentComplete 80
foreach($card in $SprintBoardCards['Inbox'])
{
    Invoke-RestMethod -Method Delete -Uri "https://api.trello.com/1/cards/$($card.id)?key=$AppKey&token=$Token" -ErrorAction Stop
}
# We delete the cards in the Assigned list. They are already in the new board.
Write-Progress -Activity "Creating new Trello sprint..." -Status "Deleting Assigned cards from $($SprintBoard.name)" -PercentComplete 85
foreach($card in $SprintBoardCards['Assigned'])
{
    Invoke-RestMethod -Method Delete -Uri "https://api.trello.com/1/cards/$($card.id)?key=$AppKey&token=$Token" -ErrorAction Stop
}
# We moving In Progress cards to To Next Sprint list, except IT work
Write-Progress -Activity "Creating new Trello sprint..." -Status "Moving In Progress cards to To Next Sprint in $($SprintBoard.name)" -PercentComplete 90
foreach($card in $SprintBoardCards['In progress'])
{
    if ($card.name -notlike "*Trabajos en sistemas internos")
    {
        $moveBody=ConvertTo-Json (@{'value'=$ToNextSprintResult.id})
        Invoke-RestMethod -Method Put -Uri "https://api.trello.com/1/cards/$($card.id)/idList?key=$AppKey&token=$Token" -ContentType "application/json" -Body $moveBody  -ErrorAction Stop
    }
}
# The same for Testing...
Write-Progress -Activity "Creating new Trello sprint..." -Status "Moving Testing cards to To Next Sprint in $($SprintBoard.name)" -PercentComplete 95
foreach($card in $SprintBoardCards['Testing'])
{
    $moveBody=ConvertTo-Json (@{'value'=$ToNextSprintResult.id})
    Invoke-RestMethod -Method Put -Uri "https://api.trello.com/1/cards/$($card.id)/idList?key=$AppKey&token=$Token" -ContentType "application/json" -Body $moveBody -ErrorAction Stop
}
# ...and document
Write-Progress -Activity "Creating new Trello sprint..." -Status "Moving Document cards to To Next Sprint in $($SprintBoard.name)" -PercentComplete 98
foreach($card in $SprintBoardCards['Document'])
{
    $moveBody=ConvertTo-Json (@{'value'=$ToNextSprintResult.id})
    Invoke-RestMethod -Method Put -Uri "https://api.trello.com/1/cards/$($card.id)/idList?key=$AppKey&token=$Token" -ContentType "application/json" -Body $moveBody -ErrorAction Stop
}
Write-Progress -Activity "Creating new Trello sprint..." -Status "All done!" -PercentComplete 100 -Completed

# Process finished! Let's report back to the user and prepare for mail sending
Write-Host ""
Write-Host -ForegroundColor Green "Process successfully completed!"
Write-Host -ForegroundColor Cyan -NoNewline "New board name: "
Write-Host $newBoard.name
Write-Host -ForegroundColor Cyan -NoNewline "URL: "
Write-Host $newBoard.shortUrl
Write-Host ""
$MailBody="<style>
body {
  font-family: 'Consolas', 'Courier-New', 'Verdana';
  font-size: 14px;
}
h1 {
    font-size: 18px;
    font-weight: bold;
}

h1.blue {
    color: blue;
}

h1.red {
    color: red;
}   

h1.green {
    color: green;
}
.copyright {
    font-size: 12px;
    font-style: italic;
}
</style>
<body>
<p>Welcome to $($newBoard.name)!</p>

<p>Please access the new kanban board using the following URL: <a href=`"$($newboard.shortUrl)`" target=`"_blank`">$($newboard.shortUrl)</a></p>
<p class=`"copyright`">This Trello report has been created using Trello PowerShell New Sprinter by Carlos Milán Figueredo</p>"

# The body has been built, so we are now handling the email sending stuff
if($SendMail)
{
    Write-Host -NoNewLine -ForegroundColor Cyan "Sending email to: "
    Write-Host $MailTo
    Write-Host -NoNewLine -ForegroundColor Cyan "Subject: "
    Write-Host $MailSubject
    Write-Host -NoNewLine -ForegroundColor Cyan "Smtp Server: "
    Write-Host -NoNewLine "$SmtpServer; "
    Write-Host -NoNewLine -ForegroundColor Cyan "Port: "
    Write-Host -NoNewLine "$SmtpPort; "
    Write-Host -NoNewline -ForegroundColor Cyan "Ssl: "
    Write-Host $SmtpSsl
    Write-Host -NoNewLine -ForegroundColor Cyan "Sending..."
    if($SmtpSsl) {
        Send-MailMessage -Subject $MailSubject -From $MailFrom -To $MailTo -SmtpServer $SmtpServer -Port $SmtpPort -Body $MailBody -BodyAsHtml -UseSsl -Credential $SmtpCred -Encoding UTF8 -ErrorAction Stop
    } else {
        Send-MailMessage -Subject $MailSubject -From $MailFrom -To $MailTo -SmtpServer $SmtpServer -Port $SmtpPort -Body $MailBody -BodyAsHtml -Credential $SmtpCred -Encoding UTF8 -ErrorAction Stop
    }
    Write-Host -ForegroundColor Green "Done"
    Write-Host ""
    Write-Host -ForegroundColor Green "Mail has been sent successfully"
}