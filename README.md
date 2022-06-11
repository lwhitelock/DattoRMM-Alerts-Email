# DattoRMM-Alerts-Email
Takes Datto RMM Alert Webhooks and sends them to an email address using the Microsoft 365 Graph API.

## Setup
### CPU and RAM information
If you wish for the script to be able to display CPU and RAM usage for the device you will need to roll out a component to datto, configure it as a monitor on your devices and set it to run as often as you would like the data to update.
This will then document RAM and CPU usage to the two custom fields you specify. (Default of 29 and 30)

### Variables
#### DattoURL
This is your Datto API URL, it can be found when you obtain an API key in Datto RMM.

#### DattoKey
This is your Datto API key for the script.

#### DattoSecretKey
This is your Datto API secret key for the script.

#### CPUUDF
This is the UDF you set to store your CPU Usage.

#### RAMUDF
This is the UDF you set to store your RAM Usage.

#### NumberOfColumns
This is the number of columns you would like to render in the email body of details sections.

#### AzureEmailClientID
Please create an application in Azure AD with the delegated mail.send permission. This is then the client ID for that application.

#### AzureEmailClientSecret
This is the secret ID for the application you created.

#### TenantID
This is the azure tenant ID for the tenant you created your application in.

#### MailFrom
Set the from email address for alert emails. This must be a mailbox that exists in your tenant, either full mailbox or a shared mailbox will work.

#### MailTo
This is the email address where you would like alerts sent to.

### Installation
To Deploy you can click this link and then configure the settings as detailed above.
[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3a%2f%2fraw.githubusercontent.com%2flwhitelock%2fDattoRMM-Alerts-Email%2fmain%2fDeployment%2fAzureDeployment.json)

#### Make a cup of tea.
It can take about 30 minutes once the deployment completes for Azure to pickup all the permissions between the different components that were deployed. After 30 minutes I would recommend restarting your function app and checking in the function options that all the key vault references show green.

## Setup in Datto RMM
To use this script you need to edit your monitors to send to a webhook.
First go to https://portal.azure.com/ and browse to your functiion app that was just deployed.
Click on Functions on the left hand side and then on 'Receive-Alert'.
At the top click on Get Function Url. This is the URL you will need to enter in Datto RMM as your webhook address.

Find the monitor you wish to edit in Datto RMM and set the URL as well as setting the body as below:
```
{
    "troubleshootingNote": "Please turn the computer off an on again to fix the issue",
    "docURL": "https://docs.yourdomain.com/alert-specific-kb-article",
    "showDeviceDetails": true,
    "showDeviceStatus": true,
    "showAlertDetails": true,
    "alertUID": "[alert_uid]",
    "alertMessage": "[alert_message]",
    "platform": "[platform]"
}
```

Save the montitor and then test it is working correctly before rolling it out for all your other monitors.

You can toggle individual details sections on and off for the monitor if they are not relevant and you can provide a link to your documentation as well as a quick troubleshooting message to help technicians with resolving issues faster.

## Troubleshooting 
If you have issues the easiest way to debug is to use VSCode with the Azure Functions extension. If you click on theAzure logo in the left and login, you can then find your function app from the list. Right click on the function and choose start streaming logs. Reset and alert so the webhook resends and look at any errors.
