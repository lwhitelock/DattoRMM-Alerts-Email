
function Get-MapColour {
    param (
        $MapList,
        $Count
    )

    $Maximum = ($MapList | measure-object).count - 1
    $Index = [array]::indexof($MapList, "$count")
    $Sixth = $Maximum / 6

    if ($count -eq 0) {
        return "rgb(34,34,34)"
    } elseif ($Index -ge 0 -and $Index -le $Sixth) {
        return "rgb(226, 230, 190)"
    } elseif ($Index -gt $Sixth -and $Index -le $Sixth * 2) {
        return "rgb(237, 223, 133)"
    } elseif ($Index -gt $Sixth * 2 -and $Index -le $Sixth * 3) {
        return "rgb(238, 203, 117)"
    } elseif ($Index -gt $Sixth * 3 -and $Index -le $Sixth * 4) {
        return "rgb(227, 174, 105)"
    } elseif ($Index -gt $Sixth * 4 -and $Index -le $Sixth * 5) {
        return "rgb(205, 137, 92)"
    } elseif ($Index -gt $Sixth * 5 -and $Index -lt $Maximum) {
        return "rgb(172, 89, 77)"
    } else {
        return "rgb(130, 34, 59)"
    }
    
}

function Get-HeatMap {
    param(
        $InputData,
        $XValues,
        $YValues
    )

    $BaseMap = [ordered]@{}
    foreach ($y in $YValues) {
        foreach ($x in $XValues) {
            $BaseMap.add("$($y)$($x)", 0)
        }
    }

    foreach ($DataToParse in $InputData) {
        $BaseMap["$($DataToParse)"] += 1
    }

    $MapValues = $BaseMap.values | Where-Object { $_ -ne 0 } | Group-Object
    $MapList = $MapValues.Name

    $HeaderRow = foreach ($x in $XValues) {
        "<th width=`"$(85/($XValues.count+1))%`" style=`"text-align:center`">$($x)</th>"
    }
    
    $HTMLRows = foreach ($y in $YValues) {
        $RowHTML = foreach ($x in $XValues) {
            '<td style="text-align:center; padding: 0; margin:0; border-collapse: collapse;"><svg height="25" width="100%" style="display:block;"><rect width="100%" height="100%" fill="' + $(Get-MapColour -MapList $MapList -Count $($BaseMap."$($y)$($x)")) + '" /></svg></td>'
        }       
        '<tr style="padding: 0; margin:0; border-spacing: 0px; border-collapse: collapse;"><td height=25px style="text-align:center; padding: 0; margin:0; border-collapse: collapse; line-height: 0px;">' + "$y</td>$RowHTML</tr>"
    }

    $Html = @"
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="padding: 0; margin:0; border-spacing: 0px; border-collapse: collapse;"><thead>
        <tr>
            <td width=15%></td>$HeaderRow
        </tr>
    </thead>
    $HTMLRows
    </table>
"@

    return $html
}


function Get-DecodedTable {
    param(
        $TableString,
        $UseValue
    )
    # "mscorsvw:48.7,system:1.3,msmpeng:0.6"
    $Parsed = $TableString -split "," | ForEach-Object {
        $Values = $_ -split ":"
        [pscustomobject]@{
            Application     = $Values[0]
            "Use $UseValue" = $Values[1]
        }
    }

    Return $Parsed

}

function Get-AlertDescription {
    param(
        $Alert
    )

    $AlertContext = $Alert.alertcontext

    switch ($AlertContext.'@class') {
        'perf_resource_usage_ctx' { $Result = "$($AlertContext.type) - $($AlertContext.percentage)" }
        'comp_script_ctx' { $Result = "$($AlertContext.Samples | convertto-html -as List -Fragment)" }
        'perf_mon_ctx' { $Result = "$($AlertContext.value)" }
        'online_offline_status_ctx' { $Result = "$($AlertContext.status)" }
        'eventlog_ctx' { $Result = "$($AlertContext.logName) - $($AlertContext.type) - $($AlertContext.code) - $($AlertContext.description)" }
        'perf_disk_usage_ctx' { $Result = "$($AlertContext.diskName) - $($AlertContext.freeSpace /1024/1024)GB free of $($AlertContext.totalVolume /1024/1024)GB" }
        'patch_ctx' { $Result = "$($AlertContext.result): $($AlertContext.info)" }
        'srvc_status_ctx' { $Result = "$($AlertContext.serviceName) - $($AlertContext.status)" }
        'antivirus_ctx' { $Result = "$($AlertContext.productName) - $($AlertContext.status)" }
        'custom_snmp_ctx' { $Result = "$($AlertContext.displayName) - $($AlertContext.currentValue)" }
        default { $Result = "Unknown Monitor Type" }
    }

    return $Result
}

function New-Email {
    param(
        $MailFrom,
        $MailTo,
        $MailSubject,
        $MailHTML
    )

    $clientID = $env:AzureEmailClientID
    $Clientsecret = $env:AzureEmailClientSecret
    $tenantID = $env:TenantID

    $MailSender = $MailFrom

    #Connect to GRAPH API
    $tokenBody = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        Client_Id     = $clientId
        Client_Secret = $clientSecret
    }
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" -Method POST -Body $tokenBody
    $headers = @{
        "Authorization" = "Bearer $($tokenResponse.access_token)"
        "Content-type"  = "application/json"
    }

    #Send Mail    
    $URLsend = "https://graph.microsoft.com/v1.0/users/$MailSender/sendMail"
    $Message = @{
        message         = @{
            subject      = $MailSubject
            body         = @{
                contentType = "HTML"
                content     = $MailHTML
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $MailTo
                    }
                })
        }
        saveToSentItems = "true"
    }

    $null = Invoke-RestMethod -Method POST -Uri $URLsend -Headers $headers -Body $($Message | convertto-json -depth 100) -ContentType "application/json"

}


function Get-HTMLBody {
    param (
        $Sections,
        $NumberOfColumns
    )

    $HTMLHeader = @"
<!-- Header HTML Start -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html>
  <head>
    <!-- Compiled with Bootstrap Email version: 1.1.3 -->
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta http-equiv="x-ua-compatible" content="ie=edge">
    <meta name="x-apple-disable-message-reformatting">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="format-detection" content="telephone=no, date=no, address=no, email=no">
    <style type="text/css">
      body,table,td{font-family:Helvetica,Arial,sans-serif !important}.ExternalClass{width:100%}.ExternalClass,.ExternalClass p,.ExternalClass span,.ExternalClass font,.ExternalClass td,.ExternalClass div{line-height:150%}a{text-decoration:none}*{color:inherit}a[x-apple-data-detectors],u+#body a,#MessageViewBody a{color:inherit;text-decoration:none;font-size:inherit;font-family:inherit;font-weight:inherit;line-height:inherit}img{-ms-interpolation-mode:bicubic}table:not([class^=s-]){font-family:Helvetica,Arial,sans-serif;mso-table-lspace:0pt;mso-table-rspace:0pt;border-spacing:0px;border-collapse:collapse}table:not([class^=s-]) td{border-spacing:0px;border-collapse:collapse}@media screen and (max-width: 1800px){.row-responsive.row{margin-right:0 !important}td.col-lg-4{display:block;width:100% !important;padding-left:0 !important;padding-right:0 !important}.max-w-96,.max-w-96>tbody>tr>td{max-width:1800px !important;width:100% !important}.w-full,.w-full>tbody>tr>td{width:100% !important}*[class*=s-lg-]>tbody>tr>td{font-size:0 !important;line-height:0 !important;height:0 !important}.s-10>tbody>tr>td{font-size:40px !important;line-height:40px !important;height:40px !important}}
    </style>
  </head>
  <body style="outline: 0; width: 100%; min-width: 100%; height: 100%; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; font-family: Helvetica, Arial, sans-serif; line-height: 24px; font-weight: normal; font-size: 16px; -moz-box-sizing: border-box; -webkit-box-sizing: border-box; box-sizing: border-box; color: #000000; margin: 0; padding: 0; border-width: 0;" bgcolor="#ffffff">
    <table class="body" valign="top" role="presentation" border="0" cellpadding="0" cellspacing="0" style="outline: 0; width: 100%; min-width: 100%; height: 100%; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; font-family: Helvetica, Arial, sans-serif; line-height: 24px; font-weight: normal; font-size: 16px; -moz-box-sizing: border-box; -webkit-box-sizing: border-box; box-sizing: border-box; color: #000000; margin: 0; padding: 0; border-width: 0;" bgcolor="#ffffff">
      <tbody style="width: 100%; max-width: 1800px; margin: 0 auto;">
        <tr>
          <td valign="top" style="line-height: 24px; font-size: 16px; margin: 0;" align="left">
            <table class="bg-black w-full" role="presentation" border="0" cellpadding="0" cellspacing="0" style="width: 100%;" bgcolor="#000000" width="100%">
              <tbody style="width: 100%; max-width: 1800px; margin: 0 auto;">
                <tr>
                  <td style="line-height: 24px; font-size: 16px; width: 100%; margin: 0;" align="left" bgcolor="#000000" width="100%">
                    <table class="container" role="presentation" border="0" cellpadding="0" cellspacing="0" style="width: 100%; max-width: 1800px;">
                      <tbody style="width: 100%; max-width: 1800px; margin: 0 auto;">
                        <tr>
                          <td align="center" style="line-height: 24px; font-size: 16px; margin: 0; padding: 0 16px;">
                            <table align="center" role="presentation" border="0" cellpadding="0" cellspacing="0" style="width: 100%; max-width: 1800px; margin: 0 auto;">
                              <tbody style="width: 100%; max-width: 1800px; margin: 0 auto;">
                                <tr>
                                  <td style="line-height: 24px; font-size: 16px; margin: 0;" align="left">
                                  <!-- Header HTML End -->
"@

    $HTMLFooter = @"
<!-- Footer HTML Start -->
                                   </td>
                                </tr>
                              </tbody>
                            </table>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
      </tbody>
    </table>
  </body>
</html>
<!-- Footer HTML End -->
"@


    $RowHeader = @"
<!-- Row Header HTML Start -->
<div class="row row-responsive" style="margin-right: -24px;">
    <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="table-layout: fixed; width: 100%;">
        <tbody>
            <tr>
            <!-- Row Header HTML End -->
"@

    $RowFooter = @"
<!-- Row Footer HTML Start -->
            </tr>
        </tbody>
    </table>
</div>
<!-- Row Footer HTML End -->
"@


    $CurrentColumn = 1
    $CalculatedWidth = 100 / $NumberOfColumns
    $SectionCount = 1

    $BlockHTML = foreach ($Section in $Sections) {

        [System.Collections.Generic.List[PSCustomObject]]$ReturnHtml = @()
        if ($currentColumn -eq 1) {
            $null = $ReturnHtml.add($RowHeader)
            Write-Host "New Row" 
        }

        Write-Host "$CurrentColumn"


        $Block = @"
    <!-- Block HTML Start -->
    <td class="col-lg-4"
        style="line-height: 24px; font-size: 16px; min-height: 1px; font-weight: normal; padding:24px; width: $CalculatedWidth%; margin: 0; background-color:#222222; border-left: 20px solid #000000; border-right: 20px solid #000000; border-top: 20px solid #000000;"
        align="left" valign="top">
        <table width="100%"class="ax-center" role="presentation" align="center" border="0" cellpadding="0" cellspacing="0"
            style="margin: 0 auto;">
            <tbody>
                <tr>
                    <td style="line-height: 24px; font-size: 16px; margin: 0;" align="left">
                        <table class="ax-center" role="presentation" align="center" border="0" cellpadding="0"
                            cellspacing="0" style="margin: 0 auto;">
                            <tbody>
                                <tr>
                                    <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" height="100%"
                                        style="background-color:#222222;">
                                        <tr>
                                            <td
                                                style="padding: 10px 10px 0px 10px; font-family: sans-serif; font-size: 15px; line-height: 20px; color: #ffffff;">
                                                <h1
                                                    style="margin: 0 0 0px; font-size: 25px; line-height: 30px; color: #ffffff; font-weight: normal;">
                                                    $($Section.Heading)</h1>
                                            </td>
                                        </tr>
                                        <!-- Block Section HTML Start -->
                                        $($Section.HTML)
                                        <!-- Block Section HTML End -->
                                    </table>
                    </td>
                </tr>
            </tbody>
        </table>
    </td>
    <!-- Block HTML End -->
"@

        $null = $ReturnHtml.add($Block)

        if (($currentColumn -eq $NumberOfColumns) -or ($SectionCount -eq $Sections.count)) {
            $null = $ReturnHtml.add($RowFooter)
            Write-Host "New Footer"
            $currentColumn = 0
        }

        $currentColumn++
        $SectionCount++
        $ReturnHtml -join ''
    }


    $HTML = $HTMLHeader + ($BlockHTML) + $HTMLFooter

    return $HTML


}

Function Get-AlertEmailBody($AlertWebhook) {
    $DattoURL = $env:DattoURL
    $DattoKey = $env:DattoKey
    $DattoSecretKey = $env:DattoSecretKey

    $CPUUDF = $env:CPUUDF
    $RAMUDF = $env:RAMUDF

    $NumberOfColumns = $env:NumberOfColumns

    $AlertTroubleshooting = $AlertWebhook.troubleshootingNote
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
        [System.Collections.Generic.List[PSCustomObject]]$Sections = @()

        $Device = Get-DrmmDevice -deviceUid $Alert.alertSourceInfo.deviceUid
        $DeviceAudit = Get-DrmmAuditDevice -deviceUid $Alert.alertSourceInfo.deviceUid

        # Build the alert details section
        Get-DRMMAlertDetailsSection -Sections $Sections -Alert $Alert -Device $Device -AlertDocumentationURL $AlertDocumentationURL -AlertTroubleshooting $AlertTroubleshooting -DattoPlatform $DattoPlatform


        ## Build the device details section if enabled.
        if ($ShowDeviceDetails -eq $True) {
            Get-DRMMDeviceDetailsSection -Sections $Sections -Device $Device
        }


        # Build the device status section if enabled
        if ($ShowDeviceStatus -eq $true) {
            Get-DRMMDeviceStatusSection -Sections $Sections -Device $Device -DeviceAudit $DeviceAudit -CPUUDF $CPUUDF -RAMUDF $RAMUDF
        }


        if ($showAlertDetails -eq $true) {
            Get-DRMMAlertHistorySection -Sections $Sections -Alert $Alert -DattoPlatform $DattoPlatform
        }

        $TicketSubject = "Alert: $($AlertTypesLookup[$Alert.alertContext.'@class']) - $($AlertMessage) on device: $($Device.hostname)"

        $HTMLBody = Get-HTMLBody -Sections $Sections -NumberOfColumns $NumberOfColumns
    
        $Email = @{
            Subject = $TicketSubject
            Body    = $HTMLBody
            Alert = $Alert
        }

        Return $Email

    } else {
        Return $Null
    }
}