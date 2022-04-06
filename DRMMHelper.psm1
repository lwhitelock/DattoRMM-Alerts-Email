
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
