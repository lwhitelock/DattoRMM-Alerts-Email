
function Get-DRMMAlertColour {
    <#
    This function returns the datto themed alert piority colours.
    #>
    param (
        $Priority
    )

    Switch ($Alert.Priority) {
        'Critical' { $Colour = ' background-color:#EC422E; color:#1C3E4C' }
        'High' { $Colour = ' background-color:#F68218; color:#1C3E4C' }
        'Moderate' { $Colour = ' background-color:#F7C210; color:#1C3E4C' }
        'Low' { $Colour = ' background-color:#2C81C8; color:#ffffff' }
        default { $Colour = 'color:#ffffff;' }
    }

    Return $Colour

}

function Get-DRMMAlertDetailsSection {
    <#
    This function returns the HTML for the alert details section.
    #>
    param(
        $Sections,
        $Alert,
        $Device,
        $AlertDocumentationURL,
        $AlertTroubleshooting,
        $DattoPlatform
    )

    if ($AlertDocumentationURL) {
        $DocLinkHTML = @"
<div style="display:inline-block; margin: 2px; max-width: 128px; min-width:100px; vertical-align:top; width:100%;"
    class="stack-column">
    <a class="button-a button-a-primary" target="_blank" href="$AlertDocumentationURL"
        style="background: #333333; border: 1px solid #000000; font-family: sans-serif; font-size: 15px; line-height: 15px; text-decoration: none; padding: 13px 17px; color: #ffffff; display: block; border-radius: 4px;">View
        Docs</a>
</div>
"@
    } else {
        $DocLinkHTML = ''
    }
    
    $Colour = Get-DRMMAlertColour -Piority $Alert.Priority

    $SectionHTML = @"
    <!-- Alert Detaills HTML Start -->
    <tr>
        <td style="padding: 20px; font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff;">
            <h1
                style="margin: 0 0 10px; font-size: 25px; line-height: 30px; font-weight: normal; $Colour">
                $($Alert.priority) Alert - $($Device.siteName) - $($Device.hostname)</h1>
            <h3>Component Monitor - [Failure Test Monitor] - Result: A Test Alert Was Created:</h3>
            <p style="margin: 0 0 10px;">$(Get-AlertDescription -Alert $Alert)
            $($Alert.diagnostics)
            </p>
            <br />
            <h3>Troubleshooting:</h3>
            <p style="margin: 0 0 10px;">$($AlertTroubleshooting)</p>
            <br />
        </td>
    </tr>

    <tr>
        <td align="center" valign="top" style="font-size:0; background-color: #222222; padding-bottom: 20px;">
            $DocLinkHTML
            <div style="display:inline-block; margin: 2px; max-width: 128px; min-width:100px; vertical-align:top; width:100%;"
                class="stack-column">
                <a class="button-a button-a-primary" target="_blank"
                    href="https://$($DattoPlatform)rmm.centrastage.net/alert/$($Alert.alertUid)"
                    style="background: #333333; border: 1px solid #000000; font-family: sans-serif; font-size: 15px; line-height: 15px; text-decoration: none; padding: 13px 17px; color: #ffffff; display: block; border-radius: 4px;">View
                    Alert</a>
            </div>
            <div style="display:inline-block; margin: 2px; max-width: 128px; min-width:100px; vertical-align:top; width:100%;"
                class="stack-column">
                <a class="button-a button-a-primary" target="_blank"
                    href="https://$($DattoPlatform)rmm.centrastage.net/device/$($Device.id)/$($Device.hostname)"
                    style="background: #333333; border: 1px solid #000000; font-family: sans-serif; font-size: 15px; line-height: 15px; text-decoration: none; padding: 13px 17px; color: #ffffff; display: block; border-radius: 4px;">View
                    Device</a>
            </div>
            <div style="display:inline-block; margin: 2px; max-width: 128px; min-width:100px; vertical-align:top; width:100%;"
                class="stack-column">
                <a class="button-a button-a-primary" target="_blank"
                    href="https://$($DattoPlatform)rmm.centrastage.net/site/$($Device.siteId)"
                    style="background: #333333; border: 1px solid #000000; font-family: sans-serif; font-size: 15px; line-height: 15px; text-decoration: none; padding: 13px 17px; color: #ffffff; display: block; border-radius: 4px;">View
                    Site</a>
            </div>
            <div style="display:inline-block; margin: 2px; max-width: 128px; min-width:100px; vertical-align:top; width:100%;"
                class="stack-column">
                <a class="button-a button-a-primary" target="_blank"
                    href="https://$($DattoPlatform).centrastage.net/csm/remote/rto/$($Device.id)"
                    style="background: #333333; border: 1px solid #000000; font-family: sans-serif; font-size: 15px; line-height: 15px; text-decoration: none; padding: 13px 17px; color: #ffffff; display: block; border-radius: 4px;">Web
                    Remote</a>
            </div>
        </td>
    </tr>

    
    <!-- Alert Details HTML End -->
"@


    $AlertDetailsSection = @{
        Heading = "Alert Details"
        HTML    = $SectionHTML
    }

    $Sections.add($AlertDetailsSection)

}



function Get-DRMMDeviceDetailsSection {
    <#
    This function returns the HTML for the device details section.
    #>
    param(
        $Device,
        $Sections
    )


    $DeviceDetailsHtml = @"
    <!-- Device Details HTML Start -->
    <tr>
        <td align="center" valign="top" style="font-size:0; padding-bottom: 10px; background-color: #222222;">
            <div style="display:inline-block; margin: 0 -1px; width:100%; min-width:200px; vertical-align:top;"
                class="stack-column">
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                    <tr>
                        <td style="padding: 10px;">
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%"
                                style="font-size: 14px; text-align: left;">
                                <tr>
                                    <td style="font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff; padding-top: 10px;"
                                        class="stack-column-center">
                                        <ul>
                                            <li>Device name: <strong>$($Device.hostname)</strong></li>
                                            <li>Site: <strong>$($Device.siteName)</strong></li>
                                            <li>User: <strong>$($Device.lastLoggedInUser)</strong></li>
                                            <li>Last Reboot: <strong>$([datetime]$origin = '1970-01-01 00:00:00';
                                                    $origin.AddMilliSeconds($Device.lastReboot))</strong></li>
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
        </td>
    </tr>
    <!-- Device Details HTML End -->
"@

    $DeviceDetailsSection = @{
        Heading = "Device Details"
        HTML    = $DeviceDetailsHtml
    }

    $Sections.add($DeviceDetailsSection)

}


function Get-DRMMDeviceStatusSection {
    <#
    This function returns the HTML for the device status section.
    #>
    param(
        $Sections,
        $Device,
        $DeviceAudit,
        $CPUUDF,
        $RAMUDF
    )

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
        <td style="padding: 20px; font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff;">
            <div style="display:inline-block; margin: 0 -1px; width:100%; min-width:200px; max-width:330px; vertical-align:top;"
                class="stack-column">
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                    <tr>
                        <td style="padding: 10px;">
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%"
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
            <div style="display:inline-block; margin: 0 -1px; width:100%; min-width:200px; max-width:330px; vertical-align:top;"
                class="stack-column">
                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                    <tr>
                        <td style="padding: 10px;">
                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%"
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
            <div style="padding: 10px 10px 0px 10px; font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff;">
                <h2>Disk Use</h3>
                    $DiskHTML
            </div>
            </td>
        </tr>
        <!-- Device Status : END -->        

"@

    $DeviceStatusSection = @{
        Heading = "Device Status"
        HTML    = $DeviceStatusHTML
    }

    $Sections.add($DeviceStatusSection)

}

function Get-DRMMAlertHistorySection {
    <#
    This function returns the HTML for the alert history section.
    #>
    param(
        $Sections,    
        $Alert,
        $DattoPlatform      
    )

    $AlertsTableStyle = '<table style="border-width: 1px; border-style: solid; border-color: white; border-collapse: collapse; table-layout: auto !important;" width=100%>'
    $AlertsTableTDStyle = '<td style = "border-width: 1px; padding: 3px; border-style: solid; border-color: white; overflow-wrap: break-word" width=auto>'

    [System.Collections.Generic.List[PSCustomObject]]$AllAlerts = @()
    $DeviceOpenAlerts = Get-DrmmDeviceOpenAlerts -deviceUid $Alert.alertSourceInfo.deviceUid
    $DeviceResolvedAlerts = Get-DrmmDeviceResolvedAlerts -deviceUid $Alert.alertSourceInfo.deviceUid

    $DeviceOpenAlerts | foreach-object { $null = $AllAlerts.add($_) }
    $DeviceResolvedAlerts | foreach-object { $null = $AllAlerts.add($_) }

    $XValues = @("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23")
    $YValues = @("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")

    $AlertDates = $AllAlerts.timestamp | Foreach-Object { [datetime]$origin = '1970-01-01 00:00:00'; $origin.AddMilliSeconds($_) }

    $ParsedDates = $AlertDates | ForEach-Object { "$($_.dayofweek)$($_.hour)" }

    $HTMLHeatmapTable = Get-Heatmap -InputData $ParsedDates -XValues $XValues -YValues $YValues


    $ParsedOpenAlerts = $DeviceOpenAlerts | select-object @{n = 'View'; e = { "<a class=`"button-a button-a-primary`" target=`"_blank`" href=`"https://$($DattoPlatform)rmm.centrastage.net/alert/$($_.alertUid)`" style=`"background: #333333; border: 1px solid #000000; font-family: sans-serif; font-size: 15px; line-height: 15px; text-decoration: none; padding: 13px 17px; color: #ffffff; display: block; border-radius: 4px;`">View</a>" } },
    @{n = 'Priority'; e = { $_.priority } },
    @{n = 'Created'; e = { $([datetime]$origin = '1970-01-01 00:00:00'; $origin.AddMilliSeconds($_.timestamp)) } },
    @{n = 'Type'; e = { $AlertTypesLookup[$_.alertContext.'@class'] } },
    @{n = 'Description'; e = { Get-AlertDescription -Alert $_ } }

    $HTMLOpenAlerts = $ParsedOpenAlerts | Sort-Object Created -desc | convertto-html -Fragment
    $HTMLParsedOpenAlerts = [System.Web.HttpUtility]::HtmlDecode(((($HTMLOpenAlerts) -replace '<table>', $AlertsTableStyle) -replace '<td>', $AlertsTableTDStyle))

    $ParsedResolvedAlerts = $DeviceResolvedAlerts | select-object @{n = 'View'; e = { "<a class=`"button-a button-a-primary`" target=`"_blank`" href=`"https://$($DattoPlatform)rmm.centrastage.net/alert/$($_.alertUid)`" style=`"background: #333333; border: 1px solid #000000; font-family: sans-serif; font-size: 15px; line-height: 15px; text-decoration: none; padding: 13px 17px; color: #ffffff; display: block; border-radius: 4px;`">View</a>" } },
    @{n = 'Priority'; e = { $_.priority } },
    @{n = 'Created'; e = { $([datetime]$origin = '1970-01-01 00:00:00'; $origin.AddMilliSeconds($_.timestamp)) } },
    @{n = 'Type'; e = { $AlertTypesLookup[$_.alertContext.'@class'] } },
    @{n = 'Description'; e = { Get-AlertDescription -Alert $_ } }

    $HTMLResolvedAlerts = $ParsedResolvedAlerts | Sort-Object Created -desc | select-object -first 10 | convertto-html -Fragment
    $HTMLParsedResolvedAlerts = [System.Web.HttpUtility]::HtmlDecode(((($HTMLResolvedAlerts) -replace '<table>', $AlertsTableStyle) -replace '<td>', $AlertsTableTDStyle))

    $AlertHistoryHTML = @"
    <!-- Alert Details : BEGIN -->
    <tr>
        <td
            style="padding: 10px 10px 0px 10px; font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff;">
            <h3>Open Alerts</h3>
    
            $($HTMLParsedOpenAlerts)
            <br />
            <h3>Recent Resolved Alerts</h3>
            $($HTMLParsedResolvedAlerts)
            <br />
        </td>
    </tr>
    <tr>
            <tr>
                <td
                    style="text-align: center; padding-bottom: 40px; font-family: sans-serif; font-size: 15px; color: #ffffff;">
                    <h2>Alert Heatmap for device</h2>
                    $($HTMLHeatmapTable)
                </td>
            </tr>
    </tr>
    <!-- Alert Details : END -->
"@
    
    $AlertHistorySection = @{
        Heading = "Alert History"
        HTML    = $AlertHistoryHTML
    }
    
    $Sections.add($AlertHistorySection)
    


}
