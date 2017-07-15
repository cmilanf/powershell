<#
.SYNOPSIS
    Generates an HTML report of the current Enterprise Team Kanban status.
    It can also sent this report via email.
.PARAMETER AppKey
    Trello application key, given by the app developer.
.PARAMETER Token
    Authentication token for the app. This is obtained by the user via
    authorization mechanism in Trello.
.PARAMETER SprintBoardId
    Short Url code of the Trello board we are going to work with
.PARAMETER PipelineBoardId
    Trello Pipeline board for reporting proposals status and customer pipe.
.PARAMETER Comments
    Custom text that will appear in head of the report body.
.PARAMETER SendMail
    If present, the script will generate and send an email with the report.
.PARAMETER MailFrom
    Email sender, it may require the -SmtpCred parameter
.PARAMETER MailTo
    Email addresses where we are going to send the report. If we are going
    to specify many, it should be separated by comma. Example:
        'carlos@domain.com,beatriz@domain.com'
.PARAMETER MailSubject
    Text of the email subject. It defaults to current date.
.PARAMETER SmtpServer
    SMTP server for sending the email report
.PARAMETER SmtpPort
    SMTP server port for innitiating the connection
.PARAMETER SmtpSsl
    If present, the report will be sent using a secure connection
.PARAMETER SmtpCred
    Credentials for authenticating in the SMTP server, if needed
.EXAMPLE
    TrelloReporter.ps1 -AppKKey 58jdf534676n98e -Token 4954sfn643643lf24a
        -SprintBoardId KJ634643 -SendEmail -SmtpCred $cred
#>

param ([Parameter(Mandatory=$true)]
    [string]$AppKey,
    [Parameter(Mandatory=$true)]
    [string]$Token,
    [Parameter(Mandatory=$true)]
    [string]$SprintBoardId,
    [Parameter(Mandatory=$false)]
    [string]$PipelineBoardId="xgtDEGZ",
    [string]$Comments,
    [switch]$SendMail=$false,
    [string]$MailFrom="Carlos Milán Figueredo",
    [string[]]$MailTo=@("ceo@domain.com";"director@domain.com";"myself@domain.com"),
    [string]$MailSubject = "$(Get-Date -Format "yyyy.MM.dd") Informe de equipo",
    [string]$SmtpServer="smtp.office365.com",
    [int]$SmtpPort=587,
    [switch]$SmtpSsl=$true,
    [System.Management.Automation.PSCredential]$SmtpCred)

Write-Host -ForegroundColor Yellow "Trello PowerShell Reporter v1.0 by Carlos Milán Figueredo - https://calnus.com"
Write-Host ""

# First, let's connect to the Trello REST API and retrieve the board cards and stuff
# We will show progress meanwhile in case user has slow connection
Write-Progress -Activity "Retrieving Trello information..." -Status "Getting Cards..." -PercentComplete 0 
$SprintBoardCardsRaw=Invoke-RestMethod "https://api.trello.com/1/board/$SprintBoardId/cards?key=$AppKey&token=$Token" -ErrorAction Stop
Write-Progress -Activity "Retrieving Trello information..." -Status "Getting Lists..." -PercentComplete 25
$SprintBoardLists=Invoke-RestMethod "https://api.trello.com/1/board/$SprintBoardId/lists?key=$AppKey&token=$Token" -ErrorAction Stop
Write-Progress -Activity "Retrieving Trello information..." -Status "Getting Labels..." -PercentComplete 50
$SprintBoardLabelsRaw=Invoke-RestMethod "https://api.trello.com/1/board/$SprintBoardId/labels?key=$AppKey&token=$Token" -ErrorAction Stop
Write-Progress -Activity "Retrieving Trello information..." -Status "Getting Pipeline Cards..." -PercentComplete 75
$PipelineBoardCards=Invoke-RestMethod "https://api.trello.com/1/board/$PipelineBoardId/cards?key=$AppKey&token=$Token" -ErrorAction Stop
Write-Progress -Activity "Retrieving Trello information..." -Status "Getting Pipeline Cards..." -PercentComplete 95
$PipelineBoardLists=Invoke-RestMethod "https://api.trello.com/1/board/$PipelineBoardId/lists?key=$AppKey&token=$Token" -ErrorAction Stop
Write-Progress -Activity "Retrieving Trello information..." -Status "All done!" -PercentComplete 100 -Completed

# These are the hashes we will be using for storing HTML output
$HtmlOutput = @{'En ejecucion'= @(); 'En espera'=@(); 'Cerrado'=@(); 'Propuestas aceptadas'=@();
  'Propuestas en progreso'=@(); 'Propuestas archivadas'=@(); 'Eventos'=@()}

# We define a HashTable of Arrays. The idea is to have something like
# SprintBoardLabels['Projects'][0] and get a card with 'Projects' label
$SprintBoardLabels = @{}
# We define a Hashtable of Hashtables of Arrays :) The idea is to have something like
# SprintBoardCards['In progress']['Projects'][0] and get a card in the 'In Progress'
# column with the 'Projects' label
$SprintBoardCards = @{}

foreach($label in $SprintBoardLabelsRaw)
{
    $SprintBoardLabels.Add($label.name, @())
}

foreach($list in $SprintBoardLists)
{
    $SprintBoardCards.Add($list.name, $SprintBoardLabels.Clone())
}

# Let's get the 3 dimensional array populated. Sadly, Trello API doesn't offer
# lists names in the /1/board/$SprintBoardId/cards, so we will have to play a bit
# with the /1/board/$SprintBoardId/lists
foreach($card in $SprintBoardCardsRaw)
{
    foreach($list in $SprintBoardLists)
    {
        if($card.idList -eq $list.id) {
            foreach($label in $SprintBoardLabels.GetEnumerator())
            {
                if ($card.labels.name -eq $label.name)
                {
                    $SprintBoardCards[$list.name][$label.name] += $card
                }
            }
        }
    }
}

# Once we got all the information organized correctly, let's prepare HTML output
foreach($card in $SprintBoardCards['In Progress']['Projects'])
{
    $HtmlOutput['En ejecucion'] += "<li>" + $card.name + "</li>"
}
foreach($card in $SprintBoardCards['In Progress']['Event'])
{
    $HtmlOutput['En ejecucion'] += "<li>" + $card.name + "</li>"
}
foreach($card in $SprintBoardCards['Testing']['Event'])
{
    $HtmlOutput['En ejecucion'] += "<li>" + $card.name + "</li>"
}
foreach($card in $SprintBoardCards['Document']['Event'])
{
    $HtmlOutput['En ejecucion'] += "<li>" + $card.name + "</li>"
}
foreach($card in $SprintBoardCards['Testing']['Projects'])
{
    $HtmlOutput['En ejecucion'] += "<li>" + $card.name + "</li>"
}
foreach($card in $SprintBoardCards['Document']['Projects'])
{
    $HtmlOutput['En ejecucion'] += "<li>" + $card.name + "</li>"
}
foreach($card in $SprintBoardCards['Assigned']['Projects'])
{
    $HtmlOutput['En espera'] += "<li>" + $card.name + "</li>"
}
foreach($card in $SprintBoardCards['Assigned']['Event'])
{
    $HtmlOutput['En espera'] += "<li>" + $card.name + "</li>"
}
foreach($card in $SprintBoardCards['Done']['Projects'])
{
    $HtmlOutput['Cerrado'] += "<li>" + $card.name + "</li>"
}
foreach($card in $SprintBoardCards['Done']['IT'])
{
    $HtmlOutput['Cerrado'] += "<li>" + $card.name + "</li>"
}
foreach($card in $SprintBoardCards['Done']['Event'])
{
    $HtmlOutput['Cerrado'] += "<li>" + $card.name + "</li>"
}
foreach($list in $PipelineBoardLists)
{
    foreach($card in $PipelineBoardCards)
    {
        if($list.id -eq $card.idList)
        {
            $HtmlOutput[$list.name]+="<li>" + $card.name + "</li>"
        }
    }
}

# With HTML output prepared, we are ready to build the actual email body
$MailBody="
<style>
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
   <p>$Comments</p>
   <h1 class=`"blue`">En ejecución</h1>
   <ul>
        <strong>$($HtmlOutput['En ejecucion'])</strong>
   </ul>
   <h1 class=`"red`">En espera</h1>
   <ul>
        $($HtmlOutput['En espera'])
   </ul>
   <h1 class=`"green`">Cerrado</h1>
   <ul>
        $($HtmlOutput['Cerrado'])
   </ul>
   <h1 class=`"green`">Propuestas aceptadas</h1>
   <ul>
        $($HtmlOutput['Propuestas aceptadas'])
   </ul>
   <h1 class=`"blue`">Propuestas en progreso</h1>
   <ul>
        $($HtmlOutput['Propuestas en progreso'])
   </ul>
   <h1 class=`"red`">Propuestas archivadas</h1>
   <ul>
        $($HtmlOutput['Propuestas archivadas'])
   </ul>
   <h1 class=`"blue`">Eventos</h1>
   <ul>
        $($HtmlOutput['Eventos'])
   </ul>
   <p class=`"copyright`">This Trello report has been created using Trello PowerShell Reporter by Carlos Milán Figueredo</p>
</body>
"
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
    Write-Host -NoNewline -ForegroundColor Cyan "Comments: "
    Write-Host $Comments
    Write-Host -NoNewLine -ForegroundColor Cyan "Sending..."
    if($SmtpSsl) {
        Send-MailMessage -Subject $MailSubject -From $MailFrom -To $MailTo -SmtpServer $SmtpServer -Port $SmtpPort -Body $MailBody -BodyAsHtml -UseSsl -Credential $SmtpCred -Encoding UTF8 -ErrorAction Stop
    } else {
        Send-MailMessage -Subject $MailSubject -From $MailFrom -To $MailTo -SmtpServer $SmtpServer -Port $SmtpPort -Body $MailBody -BodyAsHtml -Credential $SmtpCred -Encoding UTF8 -ErrorAction Stop
    }
    Write-Host -ForegroundColor Green "Done"
    Write-Host ""
    Write-Host -ForegroundColor Green "Report has been sent successfully"
} else {
    # If we are not to send the email, we are returning the HTML body
    return $MailBody
}
