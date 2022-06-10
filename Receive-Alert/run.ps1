using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

Write-Host "Processing Webhook for Alert $($Request.Body.alertUID)"

$AlertWebhook = $Request.Body

$Email = Get-AlertEmailBody -AlertWebhook $AlertWebhook

if ($Email) {
    $EmailSubject = $Email.Subject
    $HTMLBody = $Email.Body    
    
    New-Email -MailFrom $env:MailFrom -MailTo $env:MailTo -MailSubject $EmailSubject -MailHTML $HtmlBody


} else {
    Write-Host "No alert found"
}



# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = ''
    })
