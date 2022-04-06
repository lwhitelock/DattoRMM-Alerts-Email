using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

Write-Host "Processing Webhook for Alert $($Request.Body.alertUID)"

$DattoURL = $env:DattoURL
$DattoKey = $env:DattoKey
$DattoSecretKey = $env:DattoSecretKey

$CPUUDF = $env:CPUUDF
$RAMUDF = $env:RAMUDF

$AlertsTableStyle = '<table style="border-width: 1px; border-style: solid; border-color: white; border-collapse: collapse; table-layout: auto !important;" width=100%>'
$AlertsTableTDStyle = '<td style = "border-width: 1px; padding: 3px; border-style: solid; border-color: white; overflow-wrap: break-word" width=auto>'


$AlertWebhook = $Request.Body

$AlertDescription = $AlertWebhook.troubleshootingNote
$AlertDocumentationURL = $AlertWebhook.docURL
$ShowDeviceDetails = $AlertWebhook.showDeviceDetails
$ShowDeviceStatus = $AlertWebhook.showDeviceStatus
$ShowAlertDetails = $AlertWebhook.showAlertDetails
$AlertID = $AlertWebhook.alertUID
$AlertMessage = $AlertWebhook.alertMessage
$DattoPlatform = $AlertWebhook.platform


$AlertTypesLookup = @{
    perf_resource_usage_ctx   = 'Resource Monitor'
    comp_script_ctx           = 'Component Monitor'
    perf_mon_ctx              = 'Performance Monitor'
    online_offline_status_ctx = 'Offline'
    eventlog_ctx              = 'Event Log'
    perf_disk_usage_ctx       = 'Disk Usage'
    patch_ctx                 = 'Patch Monitor'
    srvc_status_ctx           = 'Service Status'
    antivirus_ctx             = 'Antivirus'
    custom_snmp_ctx           = 'SNMP'
}


$params = @{
    Url       = $DattoURL
    Key       = $DattoKey
    SecretKey = $DattoSecretKey
}

Set-DrmmApiParameters @params

$Alert = Get-DrmmAlert -alertUid $AlertID

if ($Alert) {

    [System.Collections.Generic.List[PSCustomObject]]$AllAlerts = @()

    $Device = Get-DrmmDevice -deviceUid $Alert.alertSourceInfo.deviceUid
    $DeviceAudit = Get-DrmmAuditDevice -deviceUid $Alert.alertSourceInfo.deviceUid


    # Generate Email Subject
    $EmailSubject = "Alert: $($AlertTypesLookup[$Alert.alertContext.'@class']) - $($AlertMessage) on device: $($Device.hostname) - Alert{$($Alert.alertUid)}"

    # Set the Header Colour for the alert based on pirority
    Switch ($Alert.priority) {
        'Critical' { $Colour = ' background-color:#EC422E; color:#1C3E4C' }
        'High' { $Colour = ' background-color:#F68218; color:#1C3E4C' }
        'Moderate' { $Colour = ' background-color:#F7C210; color:#1C3E4C' }
        'Low' { $Colour = ' background-color:#2C81C8; color:#ffffff' }
        default { $Colour = 'color:#ffffff;' }
    }


    # Build the Documentation Link Button
    if ($AlertDocumentationURL) {
        $DocLinkHTML = @"
<td valign="top" width="128">
                        <![endif]-->
                                    <div style="display:inline-block; margin: 2px; max-width: 128px; min-width:100px; vertical-align:top; width:100%;"
                                        class="stack-column">
                                        <a class="button-a button-a-primary" target="_blank" href="$AlertDocumentationURL"
                                            style="background: #333333; border: 1px solid #000000; font-family: sans-serif; font-size: 15px; line-height: 15px; text-decoration: none; padding: 13px 17px; color: #ffffff; display: block; border-radius: 4px;">View
                                            Docs</a>
                                    </div>
                                    <!--[if mso]>
                        </td>
"@
    } else {
        $DocLinkHTML = ''
    }


    # Build the device details section if enabled.
    if ($ShowDeviceDetails -eq $True) {
        $DeviceDetailsHTML = @"
    <!-- Device Details : BEGIN -->
    <tr>
        <td style="background-color: #222222;">
            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                <tr>
                    <td
                        style="padding: 10px 10px 0px 10px; font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff;">
                        <h1
                            style="margin: 0 0 0px; font-size: 25px; line-height: 30px; color: #ffffff; font-weight: normal;">
                            Device Details</h1>
                    </td>
                </tr>
                <tr>
                    <td align="center" valign="top"
                        style="font-size:0; padding-bottom: 10px; background-color: #222222;">
                        <!--[if mso]>
            <table role="presentation" border="0" cellspacing="0" cellpadding="0" width="660">
            <tr>
            <td valign="top" width="330">
            <![endif]-->
                        <div style="display:inline-block; margin: 0 -1px; width:100%; min-width:200px; max-width:330px; vertical-align:top;"
                            class="stack-column">
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0"
                                width="100%">
                                <tr>
                                    <td style="padding: 10px;">
                                        <table role="presentation" cellspacing="0" cellpadding="0"
                                            border="0" width="100%"
                                            style="font-size: 14px; text-align: left;">
                                            <tr>
                                                <td style="font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff; padding-top: 10px;"
                                                    class="stack-column-center">
                                                    <ul>
                                                        <li>Device name: <strong>$($Device.hostname)</strong></li>
                                                        <li>Site: <strong>$($Device.siteName)</strong></li>
                                                        <li>User: <strong>$($Device.lastLoggedInUser)</strong></li>

                                                    </ul>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </div>
                        <!--[if mso]>
            </td>
            <td valign="top" width="330">
            <![endif]-->
                        <div style="display:inline-block; margin: 0 -1px; width:100%; min-width:200px; max-width:330px; vertical-align:top;"
                            class="stack-column">
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0"
                                width="100%">
                                <tr>
                                    <td style="padding: 10px;">
                                        <table role="presentation" cellspacing="0" cellpadding="0"
                                            border="0" width="100%"
                                            style="font-size: 14px;text-align: left;">
                                            <tr>
                                                <td style="font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff; padding-top: 10px;"
                                                    class="stack-column-center">
                                                    <ul>
                                                        <li>Last Reboot: <strong>$([datetime]$origin = '1970-01-01 00:00:00'; $origin.AddMilliSeconds($Device.lastReboot))</strong></li>
                                                        <li>Internal IP: <strong>$($Device.intIpAddress)</strong></li>
                                                        <li>External IP: <strong>$($Device.extIpAddress)</strong></li>
                                                    </ul>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </div>
                        <!--[if mso]>
            </td>
            </tr>
            </table>
            <![endif]-->
                    </td>
                </tr>

            </table>
        </td>
    </tr>
    <!-- Device Details : END -->
    <tr>
    <td style="padding: 20px 0; text-align: center">
    </td>
</tr>
"@
    } else {
        $DeviceDetailsHTML = ''
    }


    # Build the device status section if enabled
    if ($ShowDeviceStatus) {

        # Generate CPU/ RAM Use Data
        $CPUData = $Device.udf."udf$CPUUDF" | convertfrom-json
        $RAMData = $Device.udf."udf$RAMUDF" | convertfrom-json

        $CPUUse = $CPUData.T
        $RAMUse = $RAMData.T

        $CPUTable = Get-DecodedTable -TableString $CPUData.D -UseValue '%' | convertto-html -Fragment
        $RAMTable = Get-DecodedTable -TableString $RAMData.D -UseValue 'GBs' | convertto-html -Fragment

        $DiskData = $DeviceAudit.logicalDisks | where-object { $_.freespace }

        # Build the HTML for Disk Usage
        $DiskRaw = foreach ($Disk in $DiskData) {
            $Total = [math]::round($Disk.size / 1024 / 1024 / 1024, 2)
            $Free = [math]::round($Disk.freespace / 1024 / 1024 / 1024, 2)
            $Used = [math]::round($Total - $Free, 2)
            $UsedPercent = [math]::round(($Used / $Total) * 100, 2)
            @"
        $($Disk.diskIdentifier) $($Used)% Used - $($Free)GB Free
        <svg width='100%' height='65px'>
            <g class='bars'>
                <rect fill='#3d5599' width='100%' height='25'></rect>;
                <rect fill='#cb4d3e' width='$($UsedPercent)%' height='25'></rect>
            </g>
        </svg>
"@
        }

        $DiskHTML = $DiskRaw -join ''
        $DeviceStatusHTML = @"
    <!-- Device Status : BEGIN -->
    <tr>
        <td style="background-color: #222222;">
            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                <tr>
                    <td
                        style="padding: 10px 10px 0px 10px; font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff;">
                        <h1
                            style="margin: 0 0 0px; font-size: 25px; line-height: 30px; color: #ffffff; font-weight: normal;">
                            Device Status</h1>
                    </td>
                </tr>
                <tr>
                    <td align="center" valign="top"
                        style="font-size:0; padding-bottom: 10px; background-color: #222222;">
                        <!--[if mso]>
            <table role="presentation" border="0" cellspacing="0" cellpadding="0" width="660">
            <tr>
            <td valign="top" width="330">
            <![endif]-->
                        <div style="display:inline-block; margin: 0 -1px; width:100%; min-width:200px; max-width:330px; vertical-align:top;"
                            class="stack-column">
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0"
                                width="100%">
                                <tr>
                                    <td style="padding: 10px;">
                                        <table role="presentation" cellspacing="0" cellpadding="0"
                                            border="0" width="100%"
                                            style="font-size: 14px; text-align: left;">
                                            <tr>
                                                <td style="font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff; padding-top: 10px;"
                                                    class="stack-column-center">
                                                    <h2>CPU Usage $($CPUUse)%</h2>
                                                    $CPUTable
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </div>
                        <!--[if mso]>
            </td>
            <td valign="top" width="330">
            <![endif]-->
                        <div style="display:inline-block; margin: 0 -1px; width:100%; min-width:200px; max-width:330px; vertical-align:top;"
                            class="stack-column">
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0"
                                width="100%">
                                <tr>
                                    <td style="padding: 10px;">
                                        <table role="presentation" cellspacing="0" cellpadding="0"
                                            border="0" width="100%"
                                            style="font-size: 14px;text-align: left;">
                                            <tr>
                                                <td style="font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff; padding-top: 10px;"
                                                    class="stack-column-center">
                                                    <h2>RAM Usage $($RAMUse)%</h2>
                                                    $RAMTable
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </div>
                        <!--[if mso]>
            </td>
            </tr>
            </table>
            <![endif]-->
                    </td>
                </tr>
                <tr>
                    <td
                        style="padding: 10px 10px 0px 10px; font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff;">
                        <h2>Disk Use</h3>
                           $DiskHTML
                    </td>
                </tr>

            </table>
        </td>
    </tr>
    <!-- Device Status : END -->
    <tr>
    <td style="padding: 20px 0; text-align: center">
    </td>
</tr>
"@
    } else {
        $DeviceStatusHTML = ''
    }


    if ($showAlertDetails -eq $true) {

        $DeviceOpenAlerts = Get-DrmmDeviceOpenAlerts -deviceUid $Alert.alertSourceInfo.deviceUid
        $DeviceResolvedAlerts = Get-DrmmDeviceResolvedAlerts -deviceUid $Alert.alertSourceInfo.deviceUid

        $DeviceOpenAlerts | foreach-object { $null = $AllAlerts.add($_) }
        $DeviceResolvedAlerts | foreach-object { $null = $AllAlerts.add($_) }

        $XValues = @("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23")
        $YValues = @("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")

        $AlertDates = $AllAlerts.timestamp | Foreach-Object { [datetime]$origin = '1970-01-01 00:00:00'; $origin.AddMilliSeconds($_) }

        $ParsedDates = $AlertDates | ForEach-Object { "$($_.dayofweek)$($_.hour)" }

        $HTMLHeatmapTable = Get-Heatmap -InputData $ParsedDates -XValues $XValues -YValues $YValues

        $ParsedOpenAlerts = $DeviceOpenAlerts | ForEach-Object {
            [PSCustomObject]@{
                View        = "<a class=`"button-a button-a-primary`" target=`"_blank`" href=`"https://$($DattoPlatform)rmm.centrastage.net/alert/$($_.alertUid)`">View</a>"
                Priority    = $_.priority
                Created     = $([datetime]$origin = '1970-01-01 00:00:00'; $origin.AddMilliSeconds($_.timestamp))
                Type        = $AlertTypesLookup[$_.alertContext.'@class']
                Description = Get-AlertDescription -Alert $_
            }
        }

        $HTMLOpenAlerts = $ParsedOpenAlerts | convertto-html -Fragment
        $HTMLParsedOpenAlerts = [System.Web.HttpUtility]::HtmlDecode(((($HTMLOpenAlerts) -replace '<table>', $AlertsTableStyle) -replace '<td>', $AlertsTableTDStyle))

        $ParsedResolvedAlerts = $DeviceResolvedAlerts | ForEach-Object { 
            [PSCustomObject]@{
                View        = "<a class=`"button-a button-a-primary`" target=`"_blank`" href=`"https://$($DattoPlatform)rmm.centrastage.net/alert/$($_.alertUid)`">View</a>"
                Priority    = $_.priority
                Created     = $([datetime]$origin = '1970-01-01 00:00:00'; $origin.AddMilliSeconds($_.timestamp))
                Type        = $AlertTypesLookup[$_.alertContext.'@class']
                Description = Get-AlertDescription -Alert $_
            }
        }

        $HTMLResolvedAlerts = $ParsedResolvedAlerts | Sort-Object Created -desc | select-object -first 10 | convertto-html -Fragment
        $HTMLParsedResolvedAlerts = [System.Web.HttpUtility]::HtmlDecode(((($HTMLResolvedAlerts) -replace '<table>', $AlertsTableStyle) -replace '<td>', $AlertsTableTDStyle))



        $AlertDetailsHTML = @"
<!-- Alert Details : BEGIN -->
<tr>
    <td style="background-color: #222222;">
        <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
            <tr>
                <td
                    style="padding: 10px 10px 0px 10px; font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff;">
                    <h1
                        style="margin: 0 0 0px; font-size: 25px; line-height: 30px; color: #ffffff; font-weight: normal;">
                        Device Alert Details</h1>
                </td>
            </tr>
            <tr>
              <td
                    style="padding: 10px 10px 0px 10px; font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff;">
                    <h3>Open Alerts</h3>

                $($HTMLParsedOpenAlerts)
<h3>Recent Resolved Alerts</h3>
                $($HTMLParsedResolvedAlerts)
                </td>
            </tr>
            <tr>
            <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%"
                style="max-width:660px; margin: auto;">
                <tr>
                    <td
                        style="text-align: center; padding-bottom: 40px; font-family: sans-serif; font-size: 15px; color: #ffffff;">
                        <h2>Alert Heatmap for device</h2>
                        $($HTMLHeatmapTable)
                    </td>
                </tr>
            </table>
            </tr>

        </table>
    </td>
</tr>
<!-- Alert Details : END -->
"@
    } else {
        $AlertDetailsHTML = ''
    }


    # Build the Full HTML Page

    $HtmlBody = @"

<!DOCTYPE html>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml">

<head>
    <meta charset="utf-8"> <!-- utf-8 works for most cases -->
    <meta name="viewport" content="width=device-width"> <!-- Forcing initial-scale shouldn't be necessary -->
    <meta http-equiv="X-UA-Compatible" content="IE=edge"> <!-- Use the latest (edge) version of IE rendering engine -->
    <meta name="x-apple-disable-message-reformatting"> <!-- Disable auto-scale in iOS 10 Mail entirely -->
    <meta name="format-detection" content="telephone=no,address=no,email=no,date=no,url=no">
    <!-- Tell iOS not to automatically link certain text strings. -->
    <meta name="color-scheme" content="light">
    <meta name="supported-color-schemes" content="light">
    <title></title> <!-- The title tag shows in email notifications, like Android 4.4. -->

    <!-- What it does: Makes background images in 72ppi Outlook render at correct size. -->
    <!--[if gte mso 9]>
    <xml>
        <o:OfficeDocumentSettings>
            <o:AllowPNG/>
            <o:PixelsPerInch>96</o:PixelsPerInch>
        </o:OfficeDocumentSettings>
    </xml>
    <![endif]-->

    <!-- Web Font / @font-face : BEGIN -->
    <!-- NOTE: If web fonts are not required, lines 23 - 41 can be safely removed. -->

    <!-- Desktop Outlook chokes on web font references and defaults to Times New Roman, so we force a safe fallback font. -->
    <!--[if mso]>
        <style>
            * {
                font-family: sans-serif !important;
            }
        </style>
    <![endif]-->

    <!-- All other clients get the webfont reference; some will render the font and others will silently fail to the fallbacks. More on that here: http://stylecampaign.com/blog/2015/02/webfont-support-in-email/ -->
    <!--[if !mso]><!-->
    <!-- insert web font reference, eg: <link href='https://fonts.googleapis.com/css?family=Roboto:400,700' rel='stylesheet' type='text/css'> -->
    <!--<![endif]-->

    <!-- Web Font / @font-face : END -->

    <!-- CSS Reset : BEGIN -->
    <style>
        /* What it does: Tells the email client that only light styles are provided but the client can transform them to dark. A duplicate of meta color-scheme meta tag above. */
        :root {
            color-scheme: light;
            supported-color-schemes: light;
        }

        /* What it does: Remove spaces around the email design added by some email clients. */
        /* Beware: It can remove the padding / margin and add a background color to the compose a reply window. */
        html,
        body {
            margin: 0 auto !important;
            padding: 0 !important;
            height: 100% !important;
            width: 100% !important;
        }

        /* What it does: Stops email clients resizing small text. */
        * {
            -ms-text-size-adjust: 100%;
            -webkit-text-size-adjust: 100%;
        }

        /* What it does: Centers email on Android 4.4 */
        div[style*="margin: 16px 0"] {
            margin: 0 !important;
        }

        /* What it does: forces Samsung Android mail clients to use the entire viewport */
        #MessageViewBody,
        #MessageWebViewDiv {
            width: 100% !important;
        }

        /* What it does: Stops Outlook from adding extra spacing to tables. */
        table,
        td {
            mso-table-lspace: 0pt !important;
            mso-table-rspace: 0pt !important;
        }

        /* What it does: Fixes webkit padding issue. */
        table {
            border-spacing: 0 !important;
            border-collapse: collapse !important;
            table-layout: fixed !important;
            margin: 0 auto !important;
        }

        /* What it does: Uses a better rendering method when resizing images in IE. */
        img {
            -ms-interpolation-mode: bicubic;
        }

        /* What it does: Prevents Windows 10 Mail from underlining links despite inline CSS. Styles for underlined links should be inline. */
        a {
            text-decoration: none;
        }

        /* What it does: A work-around for email clients meddling in triggered links. */
        a[x-apple-data-detectors],
        /* iOS */
        .unstyle-auto-detected-links a,
        .aBn {
            border-bottom: 0 !important;
            cursor: default !important;
            color: inherit !important;
            text-decoration: none !important;
            font-size: inherit !important;
            font-family: inherit !important;
            font-weight: inherit !important;
            line-height: inherit !important;
        }

        /* What it does: Prevents Gmail from changing the text color in conversation threads. */
        .im {
            color: inherit !important;
        }

        /* What it does: Prevents Gmail from displaying a download button on large, non-linked images. */
        .a6S {
            display: none !important;
            opacity: 0.01 !important;
        }

        /* If the above doesn't work, add a .g-img class to any image in question. */
        img.g-img+div {
            display: none !important;
        }

        /* What it does: Removes right gutter in Gmail iOS app: https://github.com/TedGoas/Cerberus/issues/89  */
        /* Create one of these media queries for each additional viewport size you'd like to fix */

        /* iPhone 4, 4S, 5, 5S, 5C, and 5SE */
        @media only screen and (min-device-width: 320px) and (max-device-width: 374px) {
            u~div .email-container {
                min-width: 320px !important;
            }
        }

        /* iPhone 6, 6S, 7, 8, and X */
        @media only screen and (min-device-width: 375px) and (max-device-width: 413px) {
            u~div .email-container {
                min-width: 375px !important;
            }
        }

        /* iPhone 6+, 7+, and 8+ */
        @media only screen and (min-device-width: 414px) {
            u~div .email-container {
                min-width: 414px !important;
            }
        }
    </style>
    <!-- CSS Reset : END -->

    <!-- Progressive Enhancements : BEGIN -->
    <style>
        /* What it does: Hover styles for buttons */
        .button-td,
        .button-a {
            transition: all 100ms ease-in;
        }

        .button-td-primary:hover,
        .button-a-primary:hover {
            background: #555555 !important;
            border-color: #ffffff !important;
        }

        /* Media Queries */
        @media screen and (max-width: 480px) {

            /* What it does: Forces table cells into full-width rows. */
            .stack-column,
            .stack-column-center {
                display: block !important;
                width: 100% !important;
                max-width: 100% !important;
                direction: ltr !important;
            }

            /* And center justify these ones. */
            .stack-column-center {
                text-align: center !important;
            }

            /* What it does: Generic utility class for centering. Useful for images, buttons, and nested tables. */
            .center-on-narrow {
                text-align: center !important;
                display: block !important;
                margin-left: auto !important;
                margin-right: auto !important;
                float: none !important;
            }

            table.center-on-narrow {
                display: inline-block !important;
            }

            /* What it does: Adjust typography on small screens to improve readability */
            .email-container p {
                font-size: 17px !important;
            }
        }
    </style>
    <!-- Progressive Enhancements : END -->

</head>
<!--
	The email background color (#222222) is defined in three places:
	1. body tag: for most email clients
	2. center tag: for Gmail and Inbox mobile apps and web versions of Gmail, GSuite, Inbox, Yahoo, AOL, Libero, Comcast, freenet, Mail.ru, Orange.fr
	3. mso conditional: For Windows 10 Mail
-->

<body width="100%" style="margin: 0; padding: 0 !important; mso-line-height-rule: exactly; background-color: #333333;">
    <center role="article" aria-roledescription="email" lang="en" style="width: 100%; background-color: #333333;">
        <!--[if mso | IE]>
    <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #333333;">
    <tr>
    <td>
    <![endif]-->

        <!-- Visually Hidden Preheader Text : BEGIN -->
        <div style="max-height:0; overflow:hidden; mso-hide:all;" aria-hidden="true">
            New $($Alert.priority) Alert - $($Device.siteName) - $($Device.hostname) - $($AlertMessage)
        </div>
        <!-- Visually Hidden Preheader Text : END -->

        <!-- Create white space after the desired preview text so email clients don’t pull other distracting text into the inbox preview. Extend as necessary. -->
        <!-- Preview Text Spacing Hack : BEGIN -->
        <div
            style="display: none; font-size: 1px; line-height: 1px; max-height: 0px; max-width: 0px; opacity: 0; overflow: hidden; mso-hide: all; font-family: sans-serif;">
            &zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;
        </div>
        <!-- Preview Text Spacing Hack : END -->

        <!--
            Set the email width. Defined in two places:
            1. max-width for all clients except Desktop Windows Outlook, allowing the email to squish on narrow but never go wider than 680px.
            2. MSO tags for Desktop Windows Outlook enforce a 680px width.
            Note: The Fluid and Responsive templates have a different width (600px). The hybrid grid is more "fragile", and I've found that 680px is a good width. Change with caution.
        -->
        <div style="max-width: 680px; margin: 0 auto;" class="email-container">
            <!--[if mso]>
            <table align="center" role="presentation" cellspacing="0" cellpadding="0" border="0" width="680">
            <tr>
            <td>
            <![endif]-->

            <!-- Email Body : BEGIN -->
            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="margin: auto;">
                <!-- Email Header : BEGIN -->
                <tr>
                    <td style="padding: 20px 0; text-align: center">
                    </td>
                </tr>
                <!-- Email Header : END -->

                <!-- Alert Details : BEGIN -->
                <tr>
                    <td style="background-color: #222222;">
                        <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                            <tr>
                                <td
                                    style="padding: 20px; font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff;">
                                    <h1
                                        style="margin: 0 0 10px; font-size: 25px; line-height: 30px; font-weight: normal;$Colour">
                                        $($Alert.priority) Alert - $($Device.siteName) - $($Device.hostname)</h1>
                                        <h3>$($AlertMessage):</h3>
                                        <p style="margin: 0 0 10px;">$($AlertTypesLookup[$Alert.alertContext.'@class']) - $(Get-AlertDescription -Alert $Alert) $($Alert.diagnostics)</p>
                                        <h3>Troubleshooting:</h3>
                                        <p style="margin: 0 0 10px;">$($AlertDescription)</p>
                                </td>
                            </tr>

                            <tr>
                                <td align="center" valign="top"
                                    style="font-size:0; background-color: #222222; padding-bottom: 20px;">
                                    <!--[if mso]>
                        <table role="presentation" border="0" cellspacing="0" cellpadding="0" width="660"  style="margin: auto;>
                        <tr>
                        $DocLinkHTML
                        <td valign="top" width="128">
                        <![endif]-->
                                    <div style="display:inline-block; margin: 2px; max-width: 128px; min-width:100px; vertical-align:top; width:100%;"
                                        class="stack-column">
                                        <a class="button-a button-a-primary" target="_blank" href="https://$($DattoPlatform)rmm.centrastage.net/alert/$($Alert.alertUid)"
                                            style="background: #333333; border: 1px solid #000000; font-family: sans-serif; font-size: 15px; line-height: 15px; text-decoration: none; padding: 13px 17px; color: #ffffff; display: block; border-radius: 4px;">View
                                            Alert</a>
                                    </div>
                                    <!--[if mso]>
                        </td>
                        <td valign="top" width="128">
                        <![endif]-->
                                    <div style="display:inline-block; margin: 2px; max-width: 128px; min-width:100px; vertical-align:top; width:100%;"
                                        class="stack-column">
                                        <a class="button-a button-a-primary" target="_blank"  href="https://$($DattoPlatform)rmm.centrastage.net/device/$($Device.id)/$($Device.hostname)"
                                            style="background: #333333; border: 1px solid #000000; font-family: sans-serif; font-size: 15px; line-height: 15px; text-decoration: none; padding: 13px 17px; color: #ffffff; display: block; border-radius: 4px;">View
                                            Device</a>
                                    </div>
                                    <!--[if mso]>
                        </td>
				<td valign="top" width="129">
                        <![endif]-->
                                    <div style="display:inline-block; margin: 2px; max-width: 128px; min-width:100px; vertical-align:top; width:100%;"
                                        class="stack-column">
                                        <a class="button-a button-a-primary" target="_blank"  href="https://$($DattoPlatform)rmm.centrastage.net/site/$($Device.siteId)"
                                            style="background: #333333; border: 1px solid #000000; font-family: sans-serif; font-size: 15px; line-height: 15px; text-decoration: none; padding: 13px 17px; color: #ffffff; display: block; border-radius: 4px;">View
                                            Site</a>
                                    </div>
                                    <!--[if mso]>
                        </td>
				<td valign="top" width="128">
                        <![endif]-->
                                    <div style="display:inline-block; margin: 2px; max-width: 128px; min-width:100px; vertical-align:top; width:100%;"
                                        class="stack-column">
                                        <a class="button-a button-a-primary" target="_blank"  href="https://$($DattoPlatform).centrastage.net/csm/remote/rto/$($Device.id)"
                                            style="background: #333333; border: 1px solid #000000; font-family: sans-serif; font-size: 15px; line-height: 15px; text-decoration: none; padding: 13px 17px; color: #ffffff; display: block; border-radius: 4px;">Web
                                            Remote</a>
                                    </div>
                                    <!--[if mso]>
                        </td>
                        </tr>
                        </table>
                        <![endif]-->
                                </td>
                            </tr>

                        </table>
                    </td>
                </tr>
                <!-- Alert Details : END -->

                <tr>
                    <td style="padding: 20px 0; text-align: center">
                    </td>
                </tr>

                $DeviceDetailsHTML
                $DeviceStatusHTML
                $AlertDetailsHTML


            <!-- Email Footer : BEGIN -->
            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%"
                style="max-width: 680px;">
                <tr>
                    <td
                        style="padding: 20px; font-family: sans-serif; font-size: 12px; line-height: 15px; text-align: center; color: #ffffff;">
                        Generated by Luke's Better Datto RMM Alerts <unsubscribe><a href="https://mspp.io">https://mspp.io</a></unsubscribe><br>
                        <br><br>
                    </td>
                </tr>
            </table>
            <!-- Email Footer : END -->

            <!--[if mso]>
            </td>
            </tr>
            </table>
            <![endif]-->
        </div>


        <!--[if mso | IE]>
    </td>
    </tr>
    </table>
    <![endif]-->
    </center>
</body>

</html>

"@

    New-Email -MailFrom $env:MailFrom -MailTo $env:MailTo -MailSubject $EmailSubject -MailHTML $HtmlBody

} else {
    Write-Host "No alert found"
}



# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = $body
    })
