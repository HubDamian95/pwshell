#$getlastDaysDate = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")

$ApplicationID = "<Application ID>"
$TenatDomainName = "<TenantDomainName>"
$AccessSecret = "<AccessSecret>"

$Body = @{    
Grant_Type    = "client_credentials"
Scope         = "https://graph.microsoft.com/.default"
client_Id     = $ApplicationID
Client_Secret = $AccessSecret
} 

$ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenatDomainName/oauth2/v2.0/token" `
-Method POST -Body $Body

$token = $ConnectGraph.access_token

$GrapGroupUrl = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D7')"
#$GrapGroupUrl = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(date=$getlastDaysDate)"

$exampleResult = Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -Uri $GrapGroupUrl -Method Get
$exampleResult | Out-File c:\temp\list.csv
$result = Import-Csv c:\temp\list.csv

$readyresult = $result | Where-Object {$_.'ColumnName' -like '<filterCriteriaYourelookingfor>'}
$captureResult = $readyUsers | Out-String

$body = "Example Line 1"
$body += "<p>Example Line 2</p>"
$body += "<p><h3><pre>$captureResult</pre></h3></p>"

Send-MailMessage -From 'example@domain.com' -To 'exampleDelivery@address.com' -Subject 'ExampleSubject' -SmtpServer 'useSMTPServer@domain.com' -BodyAsHTML $body