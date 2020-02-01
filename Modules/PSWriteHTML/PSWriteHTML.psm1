﻿function Compare-MultipleObjects { 
    [CmdLetBinding()]
    param([System.Collections.IList] $Objects,
        [switch] $CompareSorted,
        [switch] $FormatOutput,
        [switch] $FormatDifferences,
        [switch] $Summary,
        [string] $Splitter = ', ',
        [string[]] $Property,
        [string[]] $ExcludeProperty,
        [switch] $AllProperties,
        [switch] $SkipProperties,
        [int] $First,
        [int] $Last,
        [Array] $Replace)
    if ($null -eq $Objects -or $Objects.Count -eq 1) {
        Write-Warning "Compare-MultipleObjects - Unable to compare objects. Not enough objects to compare ($($Objects.Count))."
        return
    }
    function Compare-TwoArrays {
        [CmdLetBinding()]
        param([string] $FieldName,
            [Array] $Object1,
            [Array] $Object2,
            [Array] $Replace)
        $Result = [ordered] @{Status = $false
            Same                     = [System.Collections.Generic.List[string]]::new()
            Add                      = [System.Collections.Generic.List[string]]::new()
            Remove                   = [System.Collections.Generic.List[string]]::new()
        }
        if ($Replace) {
            foreach ($R in $Replace) {
                if (($($R.Keys[0]) -eq '') -or ($($R.Keys[0]) -eq $FieldName)) {
                    if ($null -ne $Object1) { $Object1 = $Object1 -replace $($R.Values)[0], $($R.Values)[1] }
                    if ($null -ne $Object2) { $Object2 = $Object2 -replace $($R.Values)[0], $($R.Values)[1] }
                }
            }
        }
        if ($null -eq $Object1 -and $null -eq $Object2) { $Result['Status'] = $true } elseif (($null -eq $Object1) -or ($null -eq $Object2)) {
            $Result['Status'] = $false
            foreach ($O in $Object1) { $Result['Add'].Add($O) }
            foreach ($O in $Object2) { $Result['Remove'].Add($O) }
        } else {
            $ComparedObject = Compare-Object -ReferenceObject $Object1 -DifferenceObject $Object2 -IncludeEqual
            foreach ($_ in $ComparedObject) { if ($_.SideIndicator -eq '==') { $Result['Same'].Add($_.InputObject) } elseif (($_.SideIndicator -eq '<=')) { $Result['Add'].Add($_.InputObject) } elseif (($_.SideIndicator -eq '=>')) { $Result['Remove'].Add($_.InputObject) } }
            IF ($Result['Add'].Count -eq 0 -and $Result['Remove'].Count -eq 0) { $Result['Status'] = $true } else { $Result['Status'] = $false }
        }
        $Result
    }
    if ($First -or $Last) {
        [int] $TotalCount = $First + $Last
        if ($TotalCount -gt 1) { $Objects = $Objects | Select-Object -First $First -Last $Last } else {
            Write-Warning "Compare-MultipleObjects - Unable to compare objects. Not enough objects to compare ($TotalCount)."
            return
        }
    }
    $ReturnValues = @($FirstElement = [ordered] @{ }
        $FirstElement['Name'] = 'Properties'
        if ($Summary) {
            $FirstElement['Same'] = $null
            $FirstElement['Different'] = $null
        }
        $FirstElement['Status'] = $false
        $FirstObjectProperties = Select-Properties -Objects $Objects -Property $Property -ExcludeProperty $ExcludeProperty -AllProperties:$AllProperties
        if (-not $SkipProperties) {
            if ($FormatOutput) { $FirstElement["Source"] = $FirstObjectProperties -join $Splitter } else { $FirstElement["Source"] = $FirstObjectProperties }
            [Array] $IsSame = for ($i = 1; $i -lt $Objects.Count; $i++) {
                if ($Objects[0] -is [System.Collections.IDictionary]) { [string[]] $CompareObjectProperties = $Objects[$i].Keys } else {
                    [string[]] $CompareObjectProperties = $Objects[$i].PSObject.Properties.Name
                    [string[]] $CompareObjectProperties = Select-Properties -Objects $Objects[$i] -Property $Property -ExcludeProperty $ExcludeProperty -AllProperties:$AllProperties
                }
                if ($FormatOutput) { $FirstElement["$i"] = $CompareObjectProperties -join $Splitter } else { $FirstElement["$i"] = $CompareObjectProperties }
                if ($CompareSorted) {
                    $Value1 = $FirstObjectProperties | Sort-Object
                    $Value2 = $CompareObjectProperties | Sort-Object
                } else {
                    $Value1 = $FirstObjectProperties
                    $Value2 = $CompareObjectProperties
                }
                $Status = Compare-TwoArrays -FieldName 'Properties' -Object1 $Value1 -Object2 $Value2 -Replace $Replace
                if ($FormatDifferences) {
                    $FirstElement["$i-Add"] = $Status['Add'] -join $Splitter
                    $FirstElement["$i-Remove"] = $Status['Remove'] -join $Splitter
                    $FirstElement["$i-Same"] = $Status['Same'] -join $Splitter
                } else {
                    $FirstElement["$i-Add"] = $Status['Add']
                    $FirstElement["$i-Remove"] = $Status['Remove']
                    $FirstElement["$i-Same"] = $Status['Same']
                }
                $Status
            }
            if ($IsSame.Status -notcontains $false) { $FirstElement['Status'] = $true } else { $FirstElement['Status'] = $false }
            if ($Summary) {
                [Array] $Collection = (0..($IsSame.Count - 1)).Where( { $IsSame[$_].Status -eq $true }, 'Split')
                if ($FormatDifferences) {
                    $FirstElement['Same'] = ($Collection[0] | ForEach-Object { $_ + 1 }) -join $Splitter
                    $FirstElement['Different'] = ($Collection[1] | ForEach-Object { $_ + 1 }) -join $Splitter
                } else {
                    $FirstElement['Same'] = $Collection[0] | ForEach-Object { $_ + 1 }
                    $FirstElement['Different'] = $Collection[1] | ForEach-Object { $_ + 1 }
                }
            }
            [PSCustomObject] $FirstElement
        }
        foreach ($_ in $FirstObjectProperties) {
            $EveryOtherElement = [ordered] @{ }
            $EveryOtherElement['Name'] = $_
            if ($Summary) {
                $EveryOtherElement['Same'] = $null
                $EveryOtherElement['Different'] = $null
            }
            $EveryOtherElement.Status = $false
            if ($FormatOutput) { $EveryOtherElement['Source'] = $Objects[0].$_ -join $Splitter } else { $EveryOtherElement['Source'] = $Objects[0].$_ }
            [Array] $IsSame = for ($i = 1; $i -lt $Objects.Count; $i++) {
                if ($FormatOutput) { $EveryOtherElement["$i"] = $Objects[$i].$_ -join $Splitter } else { $EveryOtherElement["$i"] = $Objects[$i].$_ }
                if ($CompareSorted) {
                    $Value1 = $Objects[0].$_ | Sort-Object
                    $Value2 = $Objects[$i].$_ | Sort-Object
                } else {
                    $Value1 = $Objects[0].$_
                    $Value2 = $Objects[$i].$_
                }
                $Status = Compare-TwoArrays -FieldName $_ -Object1 $Value1 -Object2 $Value2 -Replace $Replace
                if ($FormatDifferences) {
                    $EveryOtherElement["$i-Add"] = $Status['Add'] -join $Splitter
                    $EveryOtherElement["$i-Remove"] = $Status['Remove'] -join $Splitter
                    $EveryOtherElement["$i-Same"] = $Status['Same'] -join $Splitter
                } else {
                    $EveryOtherElement["$i-Add"] = $Status['Add']
                    $EveryOtherElement["$i-Remove"] = $Status['Remove']
                    $EveryOtherElement["$i-Same"] = $Status['Same']
                }
                $Status
            }
            if ($IsSame.Status -notcontains $false) { $EveryOtherElement['Status'] = $true } else { $EveryOtherElement['Status'] = $false }
            if ($Summary) {
                [Array] $Collection = (0..($IsSame.Count - 1)).Where( { $IsSame[$_].Status -eq $true }, 'Split')
                if ($FormatDifferences) {
                    $EveryOtherElement['Same'] = ($Collection[0] | ForEach-Object { $_ + 1 }) -join $Splitter
                    $EveryOtherElement['Different'] = ($Collection[1] | ForEach-Object { $_ + 1 }) -join $Splitter
                } else {
                    $EveryOtherElement['Same'] = $Collection[0] | ForEach-Object { $_ + 1 }
                    $EveryOtherElement['Different'] = $Collection[1] | ForEach-Object { $_ + 1 }
                }
            }
            [PSCuStomObject] $EveryOtherElement
        })
    if ($ReturnValues.Count -eq 1) { return , $ReturnValues } else { return $ReturnValues }
}
function Convert-Color { 
    <#
    .Synopsis
    This color converter gives you the hexadecimal values of your RGB colors and vice versa (RGB to HEX)
    .Description
    This color converter gives you the hexadecimal values of your RGB colors and vice versa (RGB to HEX). Use it to convert your colors and prepare your graphics and HTML web pages.
    .Parameter RBG
    Enter the Red Green Blue value comma separated. Red: 51 Green: 51 Blue: 204 for example needs to be entered as 51,51,204
    .Parameter HEX
    Enter the Hex value to be converted. Do not use the '#' symbol. (Ex: 3333CC converts to Red: 51 Green: 51 Blue: 204)
    .Example
    .\convert-color -hex FFFFFF
    Converts hex value FFFFFF to RGB

    .Example
    .\convert-color -RGB 123,200,255
    Converts Red = 123 Green = 200 Blue = 255 to Hex value

    #>
    param([Parameter(ParameterSetName = "RGB", Position = 0)]
        [ValidateScript( { $_ -match '^([01]?[0-9]?[0-9]|2[0-4][0-9]|25[0-5])$' })]
        $RGB,
        [Parameter(ParameterSetName = "HEX", Position = 0)]
        [ValidateScript( { $_ -match '[A-Fa-f0-9]{6}' })]
        [string]
        $HEX)
    switch ($PsCmdlet.ParameterSetName) {
        "RGB" {
            if ($null -eq $RGB[2]) { Write-Error "Value missing. Please enter all three values seperated by comma." }
            $red = [convert]::Tostring($RGB[0], 16)
            $green = [convert]::Tostring($RGB[1], 16)
            $blue = [convert]::Tostring($RGB[2], 16)
            if ($red.Length -eq 1) { $red = '0' + $red }
            if ($green.Length -eq 1) { $green = '0' + $green }
            if ($blue.Length -eq 1) { $blue = '0' + $blue }
            Write-Output $red$green$blue
        }
        "HEX" {
            $red = $HEX.Remove(2, 4)
            $Green = $HEX.Remove(4, 2)
            $Green = $Green.remove(0, 2)
            $Blue = $hex.Remove(0, 4)
            $Red = [convert]::ToInt32($red, 16)
            $Green = [convert]::ToInt32($green, 16)
            $Blue = [convert]::ToInt32($blue, 16)
            Write-Output $red, $Green, $blue
        }
    }
}
function Get-FileName { 
    <#
    .SYNOPSIS
    Short description

    .DESCRIPTION
    Long description

    .PARAMETER Extension
    Parameter description

    .PARAMETER Temporary
    Parameter description

    .PARAMETER TemporaryFileOnly
    Parameter description

    .EXAMPLE
    Get-FileName -Temporary
    Output: 3ymsxvav.tmp

    .EXAMPLE

    Get-FileName -Temporary
    Output: C:\Users\pklys\AppData\Local\Temp\tmpD74C.tmp

    .EXAMPLE

    Get-FileName -Temporary -Extension 'xlsx'
    Output: C:\Users\pklys\AppData\Local\Temp\tmp45B6.xlsx


    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param([string] $Extension = 'tmp',
        [switch] $Temporary,
        [switch] $TemporaryFileOnly)
    if ($Temporary) { return "$($([System.IO.Path]::GetTempFileName()).Replace('.tmp','')).$Extension" }
    if ($TemporaryFileOnly) { return "$($([System.IO.Path]::GetRandomFileName()).Split('.')[0]).$Extension" }
}
function Get-RandomStringName { 
    [cmdletbinding()]
    param([int] $Size = 31,
        [switch] $ToLower,
        [switch] $ToUpper,
        [switch] $LettersOnly)
    [string] $MyValue = @(if ($LettersOnly) { ( -join ((1..$Size) | ForEach-Object { (65..90) + (97..122) | Get-Random } | ForEach-Object { [char]$_ })) } else { ( -join ((48..57) + (97..122) | Get-Random -Count $Size | ForEach-Object { [char]$_ })) })
    if ($ToLower) { return $MyValue.ToLower() }
    if ($ToUpper) { return $MyValue.ToUpper() }
    return $MyValue
}
function Remove-EmptyValues { 
    [CmdletBinding()]
    param([System.Collections.IDictionary] $Hashtable,
        [switch] $Recursive,
        [int] $Rerun)
    foreach ($_ in [string[]] $Hashtable.Keys) { if ($Recursive) { if ($Hashtable[$_] -is [System.Collections.IDictionary]) { if ($Hashtable[$_].Count -eq 0) { $Hashtable.Remove($_) } else { Remove-EmptyValues -Hashtable $Hashtable[$_] -Recursive:$Recursive } } else { if ($null -eq $Hashtable[$_]) { $Hashtable.Remove($_) } elseif ($Hashtable[$_] -is [string] -and $Hashtable[$_] -eq '') { $Hashtable.Remove($_) } } } else { if ($null -eq $Hashtable[$_]) { $Hashtable.Remove($_) } elseif ($Hashtable[$_] -is [string] -and $Hashtable[$_] -eq '') { $Hashtable.Remove($_) } } }
    if ($Rerun) { for ($i = 0; $i -lt $Rerun; $i++) { Remove-EmptyValues -Hashtable $Hashtable -Recursive:$Recursive } }
}
function Select-Properties { 
    [CmdLetBinding()]
    param([Array] $Objects,
        [string[]] $Property,
        [string[]] $ExcludeProperty,
        [switch] $AllProperties)
    function Select-Unique {
        [CmdLetBinding()]
        param([System.Collections.IList] $Object)
        $New = $Object.ToLower() | Select-Object -Unique
        $Selected = foreach ($_ in $New) {
            $Index = $Object.ToLower().IndexOf($_)
            if ($Index -ne -1) { $Object[$Index] }
        }
        $Selected
    }
    if ($Objects.Count -eq 0) {
        Write-Warning 'Select-Properties - Unable to process. Objects count equals 0.'
        return
    }
    if ($Objects[0] -is [System.Collections.IDictionary]) {
        if ($AllProperties) {
            [Array] $All = foreach ($_ in $Objects) { $_.Keys }
            $FirstObjectProperties = Select-Unique -Object $All
        } else { $FirstObjectProperties = $Objects[0].Keys }
        if ($Property.Count -gt 0 -and $ExcludeProperty.Count -gt 0) {
            $FirstObjectProperties = foreach ($_ in $FirstObjectProperties) {
                if ($Property -contains $_ -and $ExcludeProperty -notcontains $_) {
                    $_
                    continue
                }
            }
        } elseif ($Property.Count -gt 0) {
            $FirstObjectProperties = foreach ($_ in $FirstObjectProperties) {
                if ($Property -contains $_) {
                    $_
                    continue
                }
            }
        } elseif ($ExcludeProperty.Count -gt 0) {
            $FirstObjectProperties = foreach ($_ in $FirstObjectProperties) {
                if ($ExcludeProperty -notcontains $_) {
                    $_
                    continue
                }
            }
        }
    } else {
        if ($Property.Count -gt 0 -and $ExcludeProperty.Count -gt 0) { $Objects = $Objects | Select-Object -Property $Property -ExcludeProperty $ExcludeProperty } elseif ($Property.Count -gt 0) { $Objects = $Objects | Select-Object -Property $Property } elseif ($ExcludeProperty.Count -gt 0) { $Objects = $Objects | Select-Object -Property '*' -ExcludeProperty $ExcludeProperty }
        if ($AllProperties) {
            [Array] $All = foreach ($_ in $Objects) { $_.PSObject.Properties.Name }
            $FirstObjectProperties = Select-Unique -Object $All
        } else { $FirstObjectProperties = $Objects[0].PSObject.Properties.Name }
    }
    return $FirstObjectProperties
}
function Send-Email { 
    [CmdletBinding(SupportsShouldProcess = $true)]
    param ([alias('EmailParameters')][System.Collections.IDictionary] $Email,
        [string] $Body,
        [string[]] $Attachment,
        [System.Collections.IDictionary] $InlineAttachments,
        [string] $Subject,
        [string[]] $To,
        [PSCustomObject] $Logger)
    try {
        if ($Email.EmailTo) { $EmailParameters = $Email.Clone() } else {
            $EmailParameters = @{EmailFrom  = $Email.From
                EmailTo                     = $Email.To
                EmailCC                     = $Email.CC
                EmailBCC                    = $Email.BCC
                EmailReplyTo                = $Email.ReplyTo
                EmailServer                 = $Email.Server
                EmailServerPassword         = $Email.Password
                EmailServerPasswordAsSecure = $Email.PasswordAsSecure
                EmailServerPasswordFromFile = $Email.PasswordFromFile
                EmailServerPort             = $Email.Port
                EmailServerLogin            = $Email.Login
                EmailServerEnableSSL        = $Email.EnableSsl
                EmailEncoding               = $Email.Encoding
                EmailEncodingSubject        = $Email.EncodingSubject
                EmailEncodingBody           = $Email.EncodingBody
                EmailSubject                = $Email.Subject
                EmailPriority               = $Email.Priority
                EmailDeliveryNotifications  = $Email.DeliveryNotifications
                EmailUseDefaultCredentials  = $Email.UseDefaultCredentials
            }
        }
    } catch {
        return @{Status = $False
            Error       = $($_.Exception.Message)
            SentTo      = ''
        }
    }
    $SmtpClient = [System.Net.Mail.SmtpClient]::new()
    if ($EmailParameters.EmailServer) { $SmtpClient.Host = $EmailParameters.EmailServer } else {
        return @{Status = $False
            Error       = "Email Server Host is not set."
            SentTo      = ''
        }
    }
    if ($EmailParameters.EmailServerPort) { $SmtpClient.Port = $EmailParameters.EmailServerPort } else {
        return @{Status = $False
            Error       = "Email Server Port is not set."
            SentTo      = ''
        }
    }
    if ($EmailParameters.EmailServerLogin) {
        $Credentials = Request-Credentials -UserName $EmailParameters.EmailServerLogin -Password $EmailParameters.EmailServerPassword -AsSecure:$EmailParameters.EmailServerPasswordAsSecure -FromFile:$EmailParameters.EmailServerPasswordFromFile -NetworkCredentials
        $SmtpClient.Credentials = $Credentials
    }
    if ($EmailParameters.EmailServerEnableSSL) { $SmtpClient.EnableSsl = $EmailParameters.EmailServerEnableSSL }
    $MailMessage = [System.Net.Mail.MailMessage]::new()
    $MailMessage.From = $EmailParameters.EmailFrom
    if ($To) { foreach ($T in $To) { $MailMessage.To.add($($T)) } } else { if ($EmailParameters.Emailto) { foreach ($To in $EmailParameters.Emailto) { $MailMessage.To.add($($To)) } } }
    if ($EmailParameters.EmailCC) { foreach ($CC in $EmailParameters.EmailCC) { $MailMessage.CC.add($($CC)) } }
    if ($EmailParameters.EmailBCC) { foreach ($BCC in $EmailParameters.EmailBCC) { $MailMessage.BCC.add($($BCC)) } }
    if ($EmailParameters.EmailReplyTo) { $MailMessage.ReplyTo = $EmailParameters.EmailReplyTo }
    $MailMessage.IsBodyHtml = $true
    if ($Subject -eq '') { $MailMessage.Subject = $EmailParameters.EmailSubject } else { $MailMessage.Subject = $Subject }
    $MailMessage.Priority = [System.Net.Mail.MailPriority]::$($EmailParameters.EmailPriority)
    if ($EmailParameters.EmailEncodingSubject) { $MailMessage.SubjectEncoding = [System.Text.Encoding]::$($EmailParameters.EmailEncodingSubject) } else { $MailMessage.SubjectEncoding = [System.Text.Encoding]::$($EmailParameters.EmailEncoding) }
    if ($EmailParameters.EmailEncodingBody) { $MailMessage.BodyEncoding = [System.Text.Encoding]::$($EmailParameters.EmailEncodingBody) } else { $MailMessage.BodyEncoding = [System.Text.Encoding]::$($EmailParameters.EmailEncoding) }
    if ($EmailParameters.EmailUseDefaultCredentials) { $SmtpClient.UseDefaultCredentials = $EmailParameters.EmailUseDefaultCredentials }
    if ($EmailParameters.EmailDeliveryNotifications) { $MailMessage.DeliveryNotificationOptions = $EmailParameters.EmailDeliveryNotifications }
    if ($PSBoundParameters.ContainsKey('InlineAttachments')) {
        $BodyPart = [Net.Mail.AlternateView]::CreateAlternateViewFromString($Body, 'text/html')
        $MailMessage.AlternateViews.Add($BodyPart)
        foreach ($Entry in $InlineAttachments.GetEnumerator()) {
            try {
                $FilePath = $Entry.Value
                Write-Verbose $FilePath
                if ($Entry.Value.StartsWith('http')) {
                    $FileName = $Entry.Value.Substring($Entry.Value.LastIndexOf("/") + 1)
                    $FilePath = Join-Path $env:temp $FileName
                    Invoke-WebRequest -Uri $Entry.Value -OutFile $FilePath
                }
                $ContentType = Get-MimeType -FileName $FilePath
                $InAttachment = [Net.Mail.LinkedResource]::new($FilePath, $ContentType)
                $InAttachment.ContentId = $Entry.Key
                $BodyPart.LinkedResources.Add($InAttachment)
            } catch {
                $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                Write-Error "Error inlining attachments: $ErrorMessage"
            }
        }
    } else { $MailMessage.Body = $Body }
    if ($PSBoundParameters.ContainsKey('Attachment')) {
        foreach ($Attach in $Attachment) {
            if (Test-Path -LiteralPath $Attach) {
                try {
                    $File = [Net.Mail.Attachment]::new($Attach)
                    $MailMessage.Attachments.Add($File)
                } catch {
                    $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                    if ($Logger) { $Logger.AddErrorRecord("Error attaching file $Attach`: $ErrorMessage") } else { Write-Error "Error attaching file $Attach`: $ErrorMessage" }
                }
            }
        }
    }
    try {
        $MailSentTo = "$($MailMessage.To) $($MailMessage.CC) $($MailMessage.BCC)".Trim()
        if ($pscmdlet.ShouldProcess("$MailSentTo", "Send-Email")) {
            $SmtpClient.Send($MailMessage)
            $MailMessage.Dispose()
            return @{Status = $True
                Error       = ""
                SentTo      = $MailSentTo
            }
        }
    } catch {
        $MailMessage.Dispose()
        return @{Status = $False
            Error       = $($_.Exception.Message)
            SentTo      = ""
        }
    }
}
function Stop-TimeLog { 
    [CmdletBinding()]
    param ([Parameter(ValueFromPipeline = $true)][System.Diagnostics.Stopwatch] $Time,
        [ValidateSet('OneLiner', 'Array')][string] $Option = 'OneLiner',
        [switch] $Continue)
    Begin { }
    Process { if ($Option -eq 'Array') { $TimeToExecute = "$($Time.Elapsed.Days) days", "$($Time.Elapsed.Hours) hours", "$($Time.Elapsed.Minutes) minutes", "$($Time.Elapsed.Seconds) seconds", "$($Time.Elapsed.Milliseconds) milliseconds" } else { $TimeToExecute = "$($Time.Elapsed.Days) days, $($Time.Elapsed.Hours) hours, $($Time.Elapsed.Minutes) minutes, $($Time.Elapsed.Seconds) seconds, $($Time.Elapsed.Milliseconds) milliseconds" } }
    End {
        if (-not $Continue) { $Time.Stop() }
        return $TimeToExecute
    }
}
function Get-MimeType { 
    [CmdletBinding()]
    param ([Parameter(Mandatory = $true)]
        [string] $FileName)
    $MimeMappings = @{'.jpeg' = 'image/jpeg'
        '.jpg'                = 'image/jpeg'
        '.png'                = 'image/png'
    }
    $Extension = [System.IO.Path]::GetExtension($FileName)
    $ContentType = $MimeMappings[ $Extension ]
    if ([string]::IsNullOrEmpty($ContentType)) { return New-Object System.Net.Mime.ContentType } else { return New-Object System.Net.Mime.ContentType($ContentType) }
}
function Request-Credentials { 
    [CmdletBinding()]
    param(
        [string] $UserName,
        [string] $Password,
        [switch] $AsSecure,
        [switch] $FromFile,
        [switch] $Output,
        [switch] $NetworkCredentials,
        [string] $Service
    )
    if ($FromFile) {
        if (($Password -ne '') -and (Test-Path $Password)) {
            # File is there and we are reading it into Password
            Write-Verbose "Request-Credentials - Reading password from file $Password"
            $Password = Get-Content -Path $Password
        } else {
            # File is not there or couldn't be read
            if ($Output) {
                return @{ Status = $false; Output = $Service; Extended = 'File with password unreadable.' }
            } else {
                Write-Warning "Request-Credentials - Secure password from file couldn't be read. File not readable. Terminating."
                return
            }
        }
    }
    if ($AsSecure) {
        try {
            $NewPassword = $Password | ConvertTo-SecureString -ErrorAction Stop
        } catch {
            $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
            if ($ErrorMessage -like '*Key not valid for use in specified state*') {
                if ($Output) {
                    return @{ Status = $false; Output = $Service; Extended = "Couldn't use credentials provided. Most likely using credentials from other user/session/computer." }
                } else {
                    Write-Warning -Message "Request-Credentials - Couldn't use credentials provided. Most likely using credentials from other user/session/computer."
                    return
                }
            } else {
                if ($Output) {
                    return @{ Status = $false; Output = $Service; Extended = $ErrorMessage }
                } else {
                    Write-Warning -Message "Request-Credentials - $ErrorMessage"
                    return
                }
            }
        }

    } else {
        $NewPassword = $Password
    }
    if ($UserName -and $NewPassword) {
        if ($AsSecure) {
            $Credentials = New-Object System.Management.Automation.PSCredential($Username, $NewPassword)
        } else {
            Try {
                $SecurePassword = $Password | ConvertTo-SecureString -asPlainText -Force -ErrorAction Stop
            } catch {
                $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                if ($ErrorMessage -like '*Key not valid for use in specified state*') {
                    if ($Output) {
                        return  @{ Status = $false; Output = $Service; Extended = "Couldn't use credentials provided. Most likely using credentials from other user/session/computer." }
                    } else {
                        Write-Warning -Message "Request-Credentials - Couldn't use credentials provided. Most likely using credentials from other user/session/computer."
                        return
                    }
                } else {
                    if ($Output) {
                        return @{ Status = $false; Output = $Service; Extended = $ErrorMessage }
                    } else {
                        Write-Warning -Message "Request-Credentials - $ErrorMessage"
                        return
                    }
                }
            }
            $Credentials = New-Object System.Management.Automation.PSCredential($Username, $SecurePassword)
        }
    } else {
        if ($Output) {
            return @{ Status = $false; Output = $Service; Extended = 'Username or/and Password is empty' }
        } else {
            Write-Warning -Message 'Request-Credentials - UserName or Password are empty.'
            return
        }
    }
    if ($NetworkCredentials) {
        return $Credentials.GetNetworkCredential()
    } else {
        return $Credentials
    }
}
function New-ApexChart {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options
    )
    $Script:HTMLSchema.Features.ChartsApex = $true
    [string] $ID = "ChartID-" + (Get-RandomStringName -Size 8)
    $Div = New-HTMLTag -Tag 'div' -Attributes @{ id = $ID; }
    $Script = New-HTMLTag -Tag 'script' -Value {
        # Convert Dictionary to JSON and return chart within SCRIPT tag
        # Make sure to return with additional empty string
        $JSON = $Options | ConvertTo-Json -Depth 5 | ForEach-Object { [System.Text.RegularExpressions.Regex]::Unescape($_) }
        "var options = $JSON"
        "var chart = new ApexCharts(document.querySelector('#$ID'),
            options
        );"
        "chart.render();"
    } -NewLine
    $Div
    # we need to move it to the end of the code therefore using additional vesel
    $Script:HTMLSchema.Charts.Add($Script)
}
function New-ChartInternalArea {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,

        [Array] $Data,
        [Array] $DataNames,
        [Array] $DataLegend,

        #[bool] $DataLabelsEnabled = $true,
        #[int] $DataLabelsOffsetX = -6,
        #[string] $DataLabelsFontSize = '12px',
        #[string] $DataLabelsColor,
        [ValidateSet('datetime', 'category', 'numeric')][string] $DataCategoriesType = 'category'

        #$Type
    )
    # Chart defintion type, size
    $Options.chart = @{
        type = 'area'
    }

    $Options.series = @( New-ChartInternalDataSet -Data $Data -DataNames $DataNames )

    # X AXIS - CATEGORIES
    $Options.xaxis = [ordered] @{ }
    if ($DataCategoriesType -ne '') {
        $Options.xaxis.type = $DataCategoriesType
    }
    if ($DataCategories.Count -gt 0) {
        $Options.xaxis.categories = $DataCategories
    }

}
function New-ChartInternalAxisX {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [string] $TitleText,
        [int] $Min,
        [int] $Max,
        [ValidateSet('datetime', 'category', 'numeric')][string] $Type = 'category',
        [Array] $Names
    )

    if (-not $Options.Contains('xaxis')) {
        $Options.xaxis = @{ }
    }
    if ($TitleText -ne '') {
        $Options.xaxis.title = @{ }
        $Options.xaxis.title.text = $TitleText
    }
    if ($MinValue -gt 0) {
        $Options.xaxis.min = $Min
    }
    if ($MinValue -gt 0) {
        $Options.xaxis.max = $Max
    }
    if ($Type -ne '') {
        $Options.xaxis.type = $Type
    }
    if ($Names.Count -gt 0) {
        $Options.xaxis.categories = $Names
    }
}
function New-ChartInternalAxisY {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [string] $TitleText,
        [int] $Min,
        [int] $Max,
        [bool] $Show,
        [bool] $ShowAlways,
        [ValidateSet('90', '270')][string] $TitleRotate = '90',
        [int] $TitleOffsetX = 0,
        [int] $TitleOffsetY = 0,
        [string] $TitleStyleColor = "Black",
        [int] $TitleStyleFontSize = 12,
        [string] $TitleStylefontFamily = 'Helvetica, Arial, sans-serif'
    )
    if (-not $Options.Contains('yaxis')) {
        $Options.yaxis = @{ }
    }

    #if ($Show) {
    $Options.yaxis.show = $Show
    $Options.yaxis.showAlways = $ShowAlways
    # }

    if ($TitleText -ne '') {
        $Options.yaxis.title = [ordered] @{ }
        $Options.yaxis.title.text = $TitleText
        $Options.yaxis.title.rotate = [int] $TitleRotate
        $Options.yaxis.title.offsetX = $TitleOffsetX
        $Options.yaxis.title.offsetY = $TitleOffsetY
        $Options.yaxis.title.style = [ordered] @{ }

        $Color = ConvertFrom-Color -Color $TitleStyleColor
        if ($null -ne $Color) {
            $Options.yaxis.title.style.color = $Coor
        }
        $Options.yaxis.title.style.fontSize = $TitleStyleFontSize
        $Options.yaxis.title.style.fontFamily = $TitleStylefontFamily
    }
    if ($Min -gt 0) {
        $Options.yaxis.min = $Min
    }
    if ($Min -gt 0) {
        $Options.yaxis.max = $Max
    }


}

<# We can build this
    yaxis: {
        show: true,
        showAlways: true,
        seriesName: undefined,
        opposite: false,
        reversed: false,
        logarithmic: false,
        tickAmount: 6,
        min: 6,
        max: 6,
        forceNiceScale: false,
        floating: false,
        decimalsInFloat: undefined,
        labels: {
            show: true,
            align: 'right',
            minWidth: 0,
            maxWidth: 160,
            style: {
                color: undefined,
                fontSize: '12px',
                fontFamily: 'Helvetica, Arial, sans-serif',
                cssClass: 'apexcharts-yaxis-label',
            },
            offsetX: 0,
            offsetY: 0,
            rotate: 0,
            formatter: (value) => { return val },
        },
        axisBorder: {
            show: true,
            color: '#78909C',
            offsetX: 0,
            offsetY: 0
        },
        axisTicks: {
            show: true,
            borderType: 'solid',
            color: '#78909C',
            width: 6,
            offsetX: 0,
            offsetY: 0
        },
        title: {
            text: undefined,
            rotate: -90,
            offsetX: 0,
            offsetY: 0,
            style: {
                color: undefined,
                fontSize: '12px',
                fontFamily: 'Helvetica, Arial, sans-serif',
                cssClass: 'apexcharts-yaxis-title',
            },
        },
        crosshairs: {
            show: true,
            position: 'back',
            stroke: {
                color: '#b6b6b6',
                width: 1,
                dashArray: 0,
            },
        },
        tooltip: {
            enabled: true,
            offsetX: 0,
        },

    }

#>
Function New-ChartInternalBar {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [bool] $Horizontal = $true,
        [bool] $DataLabelsEnabled = $true,
        [int] $DataLabelsOffsetX = -6,
        [string] $DataLabelsFontSize = '12px',
        [string[]] $DataLabelsColor,
        [string] $Title,
        [ValidateSet('center', 'left', 'right', 'default')][string] $TitleAlignment = 'default',
        [string] $Formatter,
        [ValidateSet('bar', 'barStacked', 'barStacked100Percent')] $Type = 'bar',
        #[string[]] $Colors,

        [switch] $Distributed,

        [Array] $Data,
        [Array] $DataNames,
        [Array] $DataLegend
    )

    if ($Type -eq 'bar') {
        $Options.chart = [ordered] @{
            type = 'bar'
        }
    } elseif ($Type -eq 'barStacked') {
        $Options.chart = [ordered] @{
            type    = 'bar'
            stacked = $true
        }
    } else {
        $Options.chart = [ordered] @{
            type      = 'bar'
            stacked   = $true
            stackType = '100%'
        }
    }

    $Options.plotOptions = @{
        bar = @{
            horizontal = $Horizontal
        }
    }
    if ($Distributed) {
        $Options.plotOptions.bar.distributed = $Distributed.IsPresent
    }
    $Options.dataLabels = [ordered] @{
        enabled = $DataLabelsEnabled
        offsetX = $DataLabelsOffsetX
        style   = @{
            fontSize = $DataLabelsFontSize
        }
    }
    if ($null -ne $DataLabelsColor) {
        $RGBColorLabel = ConvertFrom-Color -Color $DataLabelsColor
        $Options.dataLabels.style.colors = @($RGBColorLabel)
    }
    $Options.series = @(New-ChartInternalDataSet -Data $Data -DataNames $DataLegend)

    # X AXIS - CATEGORIES
    $Options.xaxis = [ordered] @{ }
    # if ($DataCategoriesType -ne '') {
    #    $Options.xaxis.type = $DataCategoriesType
    #}
    if ($DataNames.Count -gt 0) {
        $Options.xaxis.categories = $DataNames
        # Need to figure out how to conver to json and leave function without ""
        #if ($Formatter -ne '') {
        #$Options.xaxis.labels = @{
        #formatter = "function(val) { return val + `"$Formatter`" }"
        #}
        #}
    }
}
Register-ArgumentCompleter -CommandName New-ChartInternalBar -ParameterName DataLabelsColor -ScriptBlock { $Script:RGBColors.Keys }
function New-ChartInternalColors {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [string[]]$Colors
    )

    if ($Colors.Count -gt 0) {
        $RGBColor = ConvertFrom-Color -Color $Colors
        $Options.colors = @($RGBColor)
    }
}
Register-ArgumentCompleter -CommandName New-ChartInternalColors -ParameterName Colors -ScriptBlock { $Script:RGBColors.Keys }
function New-ChartInternalDataLabels {
    param(
        [System.Collections.IDictionary] $Options,
        [bool] $DataLabelsEnabled = $true,
        [int] $DataLabelsOffsetX = -6,
        [string] $DataLabelsFontSize = '12px',
        [string[]] $DataLabelsColor
    )

    $Options.dataLabels = [ordered] @{
        enabled = $DataLabelsEnabled
        offsetX = $DataLabelsOffsetX
        style   = @{
            fontSize = $DataLabelsFontSize
        }
    }
    if ($DataLabelsColor.Count -gt 0) {
        $Options.dataLabels.style.colors = @(ConvertFrom-Color -Color $DataLabelsColor)
    }
}
Register-ArgumentCompleter -CommandName New-ChartInternalDataLabels -ParameterName DataLabelsColors -ScriptBlock { $Script:RGBColors.Keys }
function New-ChartInternalDataSet {
    [CmdletBinding()]
    param(
        [Array] $Data,
        [Array] $DataNames
    )

    if ($null -ne $Data -and $null -ne $DataNames) {
        if ($Data[0] -is [System.Collections.ICollection]) {
            # If it's array of Arrays
            if ($Data[0].Count -eq $DataNames.Count) {
                for ($a = 0; $a -lt $Data.Count; $a++) {
                    [ordered] @{
                        name = $DataNames[$a]
                        data = $Data[$a]
                    }
                }
            } elseif ($Data.Count -eq $DataNames.Count) {
                for ($a = 0; $a -lt $Data.Count; $a++) {
                    [ordered] @{
                        name = $DataNames[$a]
                        data = $Data[$a]
                    }
                }
            } else {
                # rerun with just data (so it checks another if)
                New-ChartInternalDataSet -Data $Data
            }

        } else {
            if ($null -ne $DataNames) {
                # If it's just int in Array
                [ordered] @{
                    name = $DataNames
                    data = $Data
                }
            } else {
                [ordered]  @{
                    data = $Data
                }
            }
        }

    } elseif ($null -ne $Data) {
        # No names given
        if ($Data[0] -is [System.Collections.ICollection]) {
            # If it's array of Arrays
            foreach ($D in $Data) {
                [ordered] @{
                    data = $D
                }
            }
        } else {
            # If it's just int in Array
            [ordered] @{
                data = $Data
            }
        }
    } else {
        Write-Warning -Message "New-ChartInternalDataSet - No Data provided. Unabled to create dataset."
        return [ordered] @{ }
    }
}
function New-ChartInternalGradient {
    [CmdletBinding()]
    param(

    )
    $Options.fill = [ordered] @{
        type     = 'gradient'
        gradient = [ordered] @{
            shade            = 'dark'
            type             = 'horizontal'
            shadeIntensity   = 0.5
            gradientToColors = @('#ABE5A1')
            inverseColors    = $true
            opacityFrom      = 1
            opacityTo        = 1
            stops            = @(0, 100)
        }
    }
}

function New-ChartInternalGrid {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [bool] $Show,
        [string] $BorderColor,
        [int] $StrokeDash, #: 0,
        [ValidateSet('front', 'back', 'default')][string] $Position = 'default',
        [nullable[bool]] $xAxisLinesShow = $null,
        [nullable[bool]] $yAxisLinesShow = $null,
        [alias('GridColors')][string[]] $RowColors,
        [alias('GridOpacity')][double] $RowOpacity = 0.5, # valid range 0 - 1
        [string[]] $ColumnColors ,
        [double] $ColumnOpacity = 0.5, # valid range 0 - 1
        [int] $PaddingTop,
        [int] $PaddingRight,
        [int] $PaddingBottom,
        [int] $PaddingLeft
    )

    <# Build this https://apexcharts.com/docs/options/grid/
        grid: {
            show: true,
            borderColor: '#90A4AE',
            strokeDashArray: 0,
            position: 'back',
            xaxis: {
                lines: {,
                    show: false
                }
            },
            yaxis: {
                lines: {,
                    show: false
                }
            },
            row: {
                colors: undefined,
                opacity: 0.5
            },
            column: {
                colors: undefined,
                opacity: 0.5
            },
            padding: {
                top: 0,
                right: 0,
                bottom: 0,
                left: 0
            },
        }
    #>

    $Options.grid = [ordered] @{ }
    $Options.grid.Show = $Show
    if ($BorderColor) {
        $options.grid.borderColor = @(ConvertFrom-Color -Color $BorderColor)
    }
    if ($StrokeDash -gt 0) {
        $Options.grid.strokeDashArray = $StrokeDash
    }
    if ($Position -ne 'Default') {
        $Options.grid.position = $Position
    }

    if ($null -ne $xAxisLinesShow) {
        $Options.grid.xaxis = @{ }
        $Options.grid.xaxis.lines = @{ }

        $Options.grid.xaxis.lines.show = $xAxisLinesShow
    }
    if ($null -ne $yAxisLinesShow) {
        $Options.grid.yaxis = @{ }
        $Options.grid.yaxis.lines = @{ }
        $Options.grid.yaxis.lines.show = $yAxisLinesShow
    }

    if ($RowColors.Count -gt 0 -or $RowOpacity -ne 0) {
        $Options.grid.row = @{ }
        if ($RowColors.Count -gt 0) {
            $Options.grid.row.colors = @(ConvertFrom-Color -Color $RowColors)
        }
        if ($RowOpacity -ne 0) {
            $Options.grid.row.opacity = $RowOpacity
        }
    }
    if ($ColumnColors.Count -gt 0 -or $ColumnOpacity -ne 0) {
        $Options.grid.column = @{ }
        if ($ColumnColors.Count -gt 0) {
            $Options.grid.column.colors = @(ConvertFrom-Color -Color $ColumnColors)
        }
        if ($ColumnOpacity -ne 0) {
            $Options.grid.column.opacity = $ColumnOpacitys
        }
    }
    if ($PaddingTop -gt 0 -or $PaddingRight -gt 0 -or $PaddingBottom -gt 0 -or $PaddingLeft -gt 0) {
        # Padding options
        $Options.grid.padding = @{ }
        if ($PaddingTop -gt 0) {
            $Options.grid.padding.PaddingTop = $PaddingTop
        }
        if ($PaddingRight -gt 0) {
            $Options.grid.padding.PaddingRight = $PaddingRight
        }
        if ($PaddingBottom -gt 0) {
            $Options.grid.padding.PaddingBottom = $PaddingBottom
        }
        if ($PaddingLeft -gt 0) {
            $Options.grid.padding.PaddingLeft = $PaddingLeft
        }
    }
}
Register-ArgumentCompleter -CommandName New-ChartInternalGrid -ParameterName BorderColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-ChartInternalGrid -ParameterName RowColors -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-ChartInternalGrid -ParameterName ColumnColors -ScriptBlock { $Script:RGBColors.Keys }
function New-ChartInternalLegend {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [ValidateSet('top', 'topRight', 'left', 'right', 'bottom', 'default')][string] $LegendPosition = 'default'
    )
    # legend
    if ($LegendPosition -eq 'default' -or $LegendPosition -eq 'bottom') {
        # Do nothing
    } elseif ($LegendPosition -eq 'right') {
        $Options.legend = [ordered]@{
            position = 'right'
            offsetY  = 100
            height   = 230
        }
    } elseif ($LegendPosition -eq 'top') {
        $Options.legend = [ordered]@{
            position        = 'top'
            horizontalAlign = 'left'
            offsetX         = 40
        }
    } elseif ($LegendPosition -eq 'topRight') {
        $Options.legend = [ordered]@{
            position        = 'top'
            horizontalAlign = 'right'
            floating        = $true
            offsetY         = -25
            offsetX         = -5
        }
    }
}
function New-ChartInternalLine {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,

        [Array] $Data,
        [Array] $DataNames,
        #[Array] $DataLegend,

        # [bool] $DataLabelsEnabled = $true,
        #[int] $DataLabelsOffsetX = -6,
        #[string] $DataLabelsFontSize = '12px',
        # [string] $DataLabelsColor,
        [ValidateSet('datetime', 'category', 'numeric')][string] $DataCategoriesType = 'category'

        # $Type
    )
    # Chart defintion type, size
    $Options.chart = @{
        type = 'line'
    }

    $Options.series = @( New-ChartInternalDataSet -Data $Data -DataNames $DataNames )

    # X AXIS - CATEGORIES
    $Options.xaxis = [ordered] @{ }
    if ($DataCategoriesType -ne '') {
        $Options.xaxis.type = $DataCategoriesType
    }
    if ($DataCategories.Count -gt 0) {
        $Options.xaxis.categories = $DataCategories
    }

}
function New-ChartInternalMarker {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [int] $MarkerSize
    )
    if ($MarkerSize -gt 0) {
        $Options.markers = @{
            size = $MarkerSize
        }
    }
}
function New-ChartInternalPattern {
    [CmdletBinding()]
    param(

    )
    $Options.fill = [ordered]@{
        type    = 'pattern'
        opacity = 1
        pattern = [ordered]@{
            style = @('circles', 'slantedLines', 'verticalLines', 'horizontalLines')
        }
    }
}
function New-ChartInternalPie {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [Array] $Values,
        [Array] $Names,
        [string] $Type
    )
    # Chart defintion type, size
    $Options.chart = @{
        type = $Type.ToLower()
    }
    $Options.series = @($Values)
    $Options.labels = @($Names)
}
function New-ChartInternalRadial {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [Array] $Values,
        [Array] $Names,
        $Type
    )
    # Chart defintion type, size
    $Options.chart = @{
        type = 'radialBar'
    }

    if ($Type -eq '1') {
        New-ChartInternalRadialType1 -Options $Options
    } elseif ($Type -eq '2') {
        New-ChartInternalRadialType2 -Options $Options
    }

    $Options.series = @($Values)
    $Options.labels = @($Names)


}
function New-ChartInternalRadialCircleType {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [ValidateSet('FullCircleTop', 'FullCircleBottom', 'FullCircleBottomLeft', 'FullCircleLeft', 'Speedometer', 'SemiCircleGauge')] $CircleType
    )
    if ($CircleType -eq 'SemiCircleGauge') {
        $Options.plotOptions.radialBar = [ordered] @{
            startAngle = -90
            endAngle   = 90
        }
    } elseif ($CircleType -eq 'FullCircleBottom') {
        $Options.plotOptions.radialBar = [ordered] @{
            startAngle = -180
            endAngle   = 180
        }
    } elseif ($CircleType -eq 'FullCircleLeft') {
        $Options.plotOptions.radialBar = [ordered] @{
            startAngle = -90
            endAngle   = 270
        }
    } elseif ($CircleType -eq 'FullCircleBottomLeft') {
        $Options.plotOptions.radialBar = [ordered] @{
            startAngle = -135
            endAngle   = 225
        }
    } elseif ($CircleType -eq 'Speedometer') {
        $Options.plotOptions.radialBar = [ordered] @{
            startAngle = -135
            endAngle   = 135
        }
    } else {
        #FullCircleTop
    }
}
function New-ChartInternalRadialDataLabels {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [string] $LabelAverage = 'Average'
    )
    if ($LabelAverage -ne '') {
        $Options.plotOptions.radialBar.dataLabels = @{
            showOn = 'always'

            name   = @{
                # fontSize = '16px'
                # color    = 'undefined'
                #offsetY = 120
            }
            value  = @{
                #offsetY = 76
                #  fontSize  = '22px'
                #  color     = 'undefined'
                # formatter = 'function (val) { return val + "%" }'
            }

            total  = @{
                show  = $true
                label = $LabelAverage
            }

        }
    }
}

function New-ChartInternalRadialType1 {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [Array] $Values,
        [Array] $Names
    )

    $Options.plotOptions = @{
        radialBar = [ordered] @{
            #startAngle = -135
            #endAngle   = 225

            #startAngle = -135
            #endAngle   = 135


            hollow = [ordered] @{
                margin       = 0
                size         = '70%'
                background   = '#fff'
                image        = 'undefined'
                imageOffsetX = 0
                imageOffsetY = 0
                position     = 'front'
                dropShadow   = @{
                    enabled = $true
                    top     = 3
                    left    = 0
                    blur    = 4
                    opacity = 0.24
                }
            }
            track  = [ordered] @{
                background  = '#fff'
                strokeWidth = '70%'
                margin      = 0  #// margin is in pixels
                dropShadow  = [ordered] @{
                    enabled = $true
                    top     = -3
                    left    = 0
                    blur    = 4
                    opacity = 0.35
                }
            }
            <#
            dataLabels = @{
                showOn = 'always'

                name   = @{
                    # fontSize = '16px'
                    # color    = 'undefined'
                    #offsetY = 120
                }
                value  = @{
                    #offsetY = 76
                    #  fontSize  = '22px'
                    #  color     = 'undefined'
                    # formatter = 'function (val) { return val + "%" }'
                }

                total  = @{
                    show  = $false
                    label = 'Average'
                }
            }
            #>
        }
    }

    $Options.fill = [ordered] @{
        type     = 'gradient'
        gradient = [ordered] @{
            shade            = 'dark'
            type             = 'horizontal'
            shadeIntensity   = 0.5
            gradientToColors = @('#ABE5A1')
            inverseColors    = $true
            opacityFrom      = 1
            opacityTo        = 1
            stops            = @(0, 100)
        }
    }
    <# Gradient
        $Options.stroke = @{
        lineCap = 'round'
    }
    #>
    <#
    $Options.fill = @{
        type     = 'gradient'
        gradient = @{
            shade          = 'dark'
            shadeIntensity = 0.15
            inverseColors  = $false
            opacityFrom    = 1
            opacityTo      = 1
            stops          = @(0, 50, 65, 91)
        }
    }
    #>
    $Options.stroke = [ordered] @{
        dashArray = 4
    }
}


function New-ChartInternalRadialType2 {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [Array] $Values,
        [Array] $Names
    )
    $Options.plotOptions = @{
        radialBar = [ordered] @{

            #startAngle = -135
            #endAngle   = 225

            #startAngle = -135
            #endAngle   = 135


            hollow = [ordered] @{
                margin       = 0
                size         = '70%'
                background   = '#fff'
                image        = 'undefined'
                imageOffsetX = 0
                imageOffsetY = 0
                position     = 'front'
                dropShadow   = @{
                    enabled = $true
                    top     = 3
                    left    = 0
                    blur    = 4
                    opacity = 0.24
                }
            }
            <#
            track      = @{
                background  = '#fff'
                strokeWidth = '70%'
                margin      = 0  #// margin is in pixels
                dropShadow  = @{
                    enabled = $true
                    top     = -3
                    left    = 0
                    blur    = 4
                    opacity = 0.35
                }
            }
            dataLabels = @{
                showOn = 'always'

                name   = @{
                    # fontSize = '16px'
                    # color    = 'undefined'
                    offsetY = 120
                }
                value  = @{
                    offsetY = 76
                    #  fontSize  = '22px'
                    #  color     = 'undefined'
                    # formatter = 'function (val) { return val + "%" }'
                }

                total  = @{
                    show  = $false
                    label = 'Average'
                }
            }
            #>
        }
    }
}

function New-ChartInternalSize {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [nullable[int]] $Height = 350,
        [nullable[int]] $Width
    )
    if ($null -ne $Height) {
        $Options.chart.height = $Height
    }
    if ($null -ne $Width) {
        $Options.chart.width = $Width
    }
}
function New-ChartInternalSpark {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [string] $Color,
        [string] $Title,
        [string] $SubTitle,
        [int] $FontSizeTitle = 24,
        [int] $FontSizeSubtitle = 14,
        [Array] $Values
    )
    if ($Values.Count -eq 0) {
        Write-Warning 'Get-ChartSpark - Values Empty'
    }

    if ($null -ne $Color) {
        $ColorRGB = ConvertFrom-Color -Color $Color
        $Options.colors = @($ColorRGB)
    }
    $Options.chart = [ordered] @{
        type      = 'area'
        sparkline = @{
            enabled = $true
        }
    }
    $Options.stroke = @{
        curve = 'straight'
    }
    $Options.title = [ordered] @{
        text    = $Title
        offsetX = 0
        style   = @{
            fontSize = "$($FontSizeTitle)px"
            cssClass = 'apexcharts-yaxis-title'
        }
    }
    $Options.subtitle = [ordered] @{
        text    = $SubTitle
        offsetX = 0
        style   = @{
            fontSize = "$($FontSizeSubtitle)px"
            cssClass = 'apexcharts-yaxis-title'
        }
    }
    $Options.yaxis = @{
        min = 0
    }
    $Options.fill = @{
        opacity = 0.3
    }
    $Options.series = @(
        # Checks if it's multiple array passed or just one. If one it will draw one line, if more then one it will draw line per each array
        if ($Values[0] -is [Array]) {
            foreach ($Value in $Values) {
                @{
                    data = @($Value)
                }
            }
        } else {
            @{
                data = @($Values)
            }
        }
    )
}

Register-ArgumentCompleter -CommandName New-ChartInternalSpark -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
function New-ChartInternalStrokeDefinition {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [bool] $LineShow = $true,
        [ValidateSet('straight', 'smooth', 'stepline')][string[]] $LineCurve,
        [int[]] $LineWidth,
        [ValidateSet('butt', 'square', 'round')][string[]] $LineCap,
        [string[]] $LineColor,
        [int[]] $LineDash
    )
    # LINE Definition
    $Options.stroke = [ordered] @{
        show = $LineShow
    }
    if ($LineCurve.Count -gt 0) {
        $Options.stroke.curve = $LineCurve
    }
    if ($LineWidth.Count -gt 0) {
        $Options.stroke.width = $LineWidth
    }
    if ($LineColor.Count -gt 0) {
        $Options.stroke.colors = @(ConvertFrom-Color -Color $LineColor)
    }
    if ($LineCap.Count -gt 0) {
        $Options.stroke.lineCap = $LineCap
    }
    if ($LineDash.Count -gt 0) {
        $Options.stroke.dashArray = $LineDash
    }
}
Register-ArgumentCompleter -CommandName New-ChartInternalStrokeDefinition -ParameterName LineColor -ScriptBlock { $Script:RGBColors.Keys }
<#
  theme: {
      mode: 'light',
      palette: 'palette1',
      monochrome: {
          enabled: false,
          color: '#255aee',
          shadeTo: 'light',
          shadeIntensity: 0.65
      },
  }
#>

function New-ChartInternalTheme {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [ValidateSet('light', 'dark')][string] $Mode,
        [ValidateSet(
            'palette1',
            'palette2',
            'palette3',
            'palette4',
            'palette5',
            'palette6',
            'palette7',
            'palette8',
            'palette9',
            'palette10'
        )
        ][string] $Palette = 'palette1',
        [switch] $Monochrome,
        [string] $Color = "DodgerBlue",
        [ValidateSet('light', 'dark')][string] $ShadeTo = 'light',
        [double] $ShadeIntensity = 0.65
    )

    $RGBColor = ConvertFrom-Color -Color $Color

    $Options.theme = [ordered] @{
        mode       = $Mode
        palette    = $Palette
        monochrome = [ordered] @{
            enabled        = $Monochrome.IsPresent
            color          = $RGBColor
            shadeTo        = $ShadeTo
            shadeIntensity = $ShadeIntensity
        }
    }
}

Register-ArgumentCompleter -CommandName New-ChartInternalTheme -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
function New-ChartInternalTitle {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [string] $Title,
        [ValidateSet('center', 'left', 'right', 'default')][string] $TitleAlignment = 'default'
    )
    # title
    $Options.title = [ordered] @{ }
    if ($TitleText -ne '') {
        $Options.title.text = $Title
    }
    if ($TitleAlignment -ne 'default') {
        $Options.title.align = $TitleAlignment
    }
}
function New-ChartInternalToolbar {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [bool] $Show = $false,
        [bool] $Download = $false,
        [bool] $Selection = $false,
        [bool] $Zoom = $false,
        [bool] $ZoomIn = $false,
        [bool] $ZoomOut = $false,
        [bool] $Pan = $false,
        [bool] $Reset = $false,
        [ValidateSet('zoom', 'selection', 'pan')][string] $AutoSelected = 'zoom'
    )
    $Options.chart.toolbar = [ordered] @{
        show         = $show
        tools        = [ordered] @{
            download  = $Download
            selection = $Selection
            zoom      = $Zoom
            zoomin    = $ZoomIn
            zoomout   = $ZoomOut
            pan       = $Pan
            reset     = $Reset
        }
        autoSelected = $AutoSelected
    }
}
function New-ChartInternalZoom {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Options,
        [switch] $Enabled
    )
    if ($Enabled) {
        $Options.chart.zoom = @{
            type    = 'x'
            enabled = $Enabled.IsPresent
        }
    }
}
function New-ChartSpark {
    [alias('ChartSpark')]
    [CmdletBinding()]
    param(
        [string] $Name,
        [object] $Value,
        [string] $Color
    )

    [PSCustomObject] @{
        ObjectType = 'Spark'
        Name       = $Name
        Value      = $Value
        Color      = $Color
    }
}

Register-ArgumentCompleter -CommandName New-ChartSpark -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
function New-HTMLChartArea {
    [CmdletBinding()]
    param(
        [nullable[int]] $Height = 350,

        [bool] $DataLabelsEnabled = $true,
        [int] $DataLabelsOffsetX = -6,
        [string] $DataLabelsFontSize = '12px',
        [string[]] $DataLabelsColor,
        [ValidateSet('datetime', 'category', 'numeric')][string] $DataCategoriesType = 'category',

        [ValidateSet('straight', 'smooth', 'stepline')] $LineCurve = 'straight',
        [int] $LineWidth,
        [string[]] $LineColor,

        [string[]] $GridColors,
        [double] $GridOpacity,

        [ValidateSet('top', 'topRight', 'left', 'right', 'bottom', 'default')][string] $LegendPosition = 'default',

        [string] $TitleX,
        [string] $TitleY,

        [int] $MarkerSize,

        [Array] $Data,
        [Array] $DataNames,
        [Array] $DataLegend,

        [switch] $Zoom,



        [string] $Title,
        [ValidateSet('center', 'left', 'right', 'default')][string] $TitleAlignment = 'default',
        [switch] $PatternedColors,
        [switch] $GradientColors,
        [System.Collections.IDictionary] $GridOptions,
        [System.Collections.IDictionary] $Toolbar,
        [System.Collections.IDictionary] $Theme
    )

    $Options = [ordered] @{ }
    New-ChartInternalArea -Options $Options -Data $Data -DataNames $DataNames

    New-ChartInternalStrokeDefinition -Options $Options `
        -LineShow $true `
        -LineCurve $LineCurve `
        -LineWidth $LineWidth `
        -LineColor $LineColor

    New-ChartInternalDataLabels -Options $Options `
        -DataLabelsEnabled $DataLabelsEnabled `
        -DataLabelsOffsetX $DataLabelsOffsetX `
        -DataLabelsFontSize $DataLabelsFontSize `
        -DataLabelsColor $DataLabelsColor

    New-ChartInternalAxisX -Options $Options `
        -Title $TitleX `
        -DataCategoriesType $DataCategoriesType `
        -DataCategories $DataLegend

    New-ChartInternalAxisY -Options $Options -Title $TitleY
    New-ChartInternalMarker -Options $Options -MarkerSize $MarkerSize
    New-ChartInternalZoom -Options $Options -Enabled:$Zoom
    New-ChartInternalLegend -Options $Options -LegendPosition $LegendPosition


    # Default for all charts
    if ($PatternedColors) { New-ChartInternalPattern }
    if ($GradientColors) { New-ChartInternalGradient }
    New-ChartInternalTitle -Options $Options -Title $Title -TitleAlignment $TitleAlignment
    New-ChartInternalSize -Options $Options -Height $Height -Width $Width
    if ($GridOptions) { New-ChartInternalGrid -Options $Options @GridOptions }
    if ($Theme) { New-ChartInternalTheme -Options $Options @Theme }
    if ($Toolbar) { New-ChartInternalToolbar -Options $Options @Toolbar -Show $true }
    New-ApexChart -Options $Options
}
Register-ArgumentCompleter -CommandName New-HTMLChartArea -ParameterName DataLabelsColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLChartArea -ParameterName LineColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLChartArea -ParameterName GridColors -ScriptBlock { $Script:RGBColors.Keys }
function New-HTMLChartBar {
    [CmdletBinding()]
    param(
        [nullable[int]] $Height = 350,
        [nullable[int]] $Width,
        [ValidateSet('bar', 'barStacked', 'barStacked100Percent')] $Type = 'bar',
        [string[]] $Colors,

        [bool] $Horizontal = $true,
        [bool] $DataLabelsEnabled = $true,
        [int] $DataLabelsOffsetX = -6,
        [string] $DataLabelsFontSize = '12px',
        [string] $DataLabelsColor,

        [switch] $Distributed,

        [ValidateSet('top', 'topRight', 'left', 'right', 'bottom', 'default')][string] $LegendPosition = 'default',

        [Array] $Data,
        [Array] $DataNames,
        [Array] $DataLegend,



        [string] $Title,
        [ValidateSet('center', 'left', 'right', 'default')][string] $TitleAlignment = 'default',
        [switch] $PatternedColors,
        [switch] $GradientColors,
        [System.Collections.IDictionary] $GridOptions,
        [System.Collections.IDictionary] $Toolbar,
        [System.Collections.IDictionary] $Theme
    )

    $Options = [ordered] @{ }
    New-ChartInternalBar -Options $Options -Horizontal $Horizontal -DataLabelsEnabled $DataLabelsEnabled `
        -DataLabelsOffsetX $DataLabelsOffsetX -DataLabelsFontSize $DataLabelsFontSize -DataLabelsColor $DataLabelsColor `
        -Data $Data -DataNames $DataNames -DataLegend $DataLegend -Title $Title -TitleAlignment $TitleAlignment `
        -Type $Type -Distributed:$Distributed

    New-ChartInternalColors -Options $Options -Colors $Colors
    New-ChartInternalLegend -Options $Options -LegendPosition $LegendPosition


    # Default for all charts
    if ($PatternedColors) { New-ChartInternalPattern }
    if ($GradientColors) { New-ChartInternalGradient }
    New-ChartInternalTitle -Options $Options -Title $Title -TitleAlignment $TitleAlignment
    New-ChartInternalSize -Options $Options -Height $Height -Width $Width
    if ($GridOptions) { New-ChartInternalGrid -Options $Options @GridOptions }
    if ($Theme) { New-ChartInternalTheme -Options $Options @Theme }
    if ($Toolbar) { New-ChartInternalToolbar -Options $Options @Toolbar -Show $true }
    New-ApexChart -Options $Options
}

Register-ArgumentCompleter -CommandName New-HTMLChartBar -ParameterName Colors -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLChartBar -ParameterName DataLabelsColor -ScriptBlock { $Script:RGBColors.Keys }
function New-HTMLChartLine {
    [CmdletBinding()]
    param(
        [nullable[int]] $Height = 350,
        [nullable[int]] $Width,

        [bool] $DataLabelsEnabled = $true,
        [int] $DataLabelsOffsetX = -6,
        [string] $DataLabelsFontSize = '12px',
        [string[]] $DataLabelsColor,
        # [ValidateSet('datetime', 'category', 'numeric')][string] $DataCategoriesType = 'category',

        [ValidateSet('straight', 'smooth', 'stepline')][string[]] $LineCurve,
        [int[]] $LineWidth,
        [string[]] $LineColor,
        [int[]] $LineDash,
        [ValidateSet('butt', 'square', 'round')][string[]] $LineCap,

        #[string[]] $GridColors,
        #[double] $GridOpacity,

        [ValidateSet('top', 'topRight', 'left', 'right', 'bottom', 'default')][string] $LegendPosition = 'default',

        #[string] $TitleX,
        #[string] $TitleY,

        [int] $MarkerSize,

        [Array] $Data,
        [Array] $DataNames,
        #[Array] $DataLegend,
        [System.Collections.IDictionary] $ChartAxisX,
        [System.Collections.IDictionary] $ChartAxisY,



        [string] $Title,
        [ValidateSet('center', 'left', 'right', 'default')][string] $TitleAlignment = 'default',
        [switch] $PatternedColors,
        [switch] $GradientColors,
        [System.Collections.IDictionary] $GridOptions,
        [System.Collections.IDictionary] $Toolbar,
        [System.Collections.IDictionary] $Theme
    )

    $Options = [ordered] @{ }

    New-ChartInternalLine -Options $Options -Data $Data -DataNames $DataNames

    if ($LineCurve.Count -eq 0 -or ($LineCurve.Count -ne $DataNames.Count)) {
        $LineCurve = for ($i = $LineCurve.Count; $i -le $DataNames.Count; $i++) {
            'straight'
        }
    }

    if ($LineCap.Count -eq 0 -or ($LineCap.Count -ne $DataNames.Count)) {
        $LineCap = for ($i = $LineCap.Count; $i -le $DataNames.Count; $i++) {
            'butt'
        }
    }
    if ($LineDash.Count -eq 0) {

    }

    New-ChartInternalStrokeDefinition -Options $Options `
        -LineShow $true `
        -LineCurve $LineCurve `
        -LineWidth $LineWidth `
        -LineColor $LineColor `
        -LineCap $LineCap `
        -LineDash $LineDash
    # line colors (stroke colors ) doesn't cover legend - we need to make sure it's the same even thou lines are already colored
    New-ChartInternalColors -Options $Options -Colors $LineColor
    New-ChartInternalDataLabels -Options $Options `
        -DataLabelsEnabled $DataLabelsEnabled `
        -DataLabelsOffsetX $DataLabelsOffsetX `
        -DataLabelsFontSize $DataLabelsFontSize `
        -DataLabelsColor $DataLabelsColor
    if ($ChartAxisX) {
        New-ChartInternalAxisX -Options $Options @ChartAxisX
    }
    if ($ChartAxisY) {
        New-ChartInternalAxisY -Options $Options @ChartAxisY
    }
    New-ChartInternalMarker -Options $Options -MarkerSize $MarkerSize
    New-ChartInternalLegend -Options $Options -LegendPosition $LegendPosition



    # Default for all charts
    if ($PatternedColors) { New-ChartInternalPattern }
    if ($GradientColors) { New-ChartInternalGradient }
    New-ChartInternalTitle -Options $Options -Title $Title -TitleAlignment $TitleAlignment
    New-ChartInternalSize -Options $Options -Height $Height -Width $Width
    if ($GridOptions) { New-ChartInternalGrid -Options $Options @GridOptions }
    if ($Theme) { New-ChartInternalTheme -Options $Options @Theme }
    if ($Toolbar) { New-ChartInternalToolbar -Options $Options @Toolbar -Show $true }
    New-ApexChart -Options $Options
}

Register-ArgumentCompleter -CommandName New-HTMLChartLine -ParameterName DataLabelsColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLChartLine -ParameterName LineColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLChartLine -ParameterName GridColors -ScriptBlock { $Script:RGBColors.Keys }
function New-HTMLChartPie {
    [CmdletBinding()]
    param(
        [string] $Type,
        [nullable[int]] $Height = 350,
        [nullable[int]] $Width,


        [bool] $DataLabelsEnabled = $true,
        [int] $DataLabelsOffsetX = -6,
        [string] $DataLabelsFontSize = '12px',
        [string[]] $DataLabelsColor,
        [Array] $Data,
        [Array] $DataNames,


        [ValidateSet('top', 'topRight', 'left', 'right', 'bottom', 'default')][string] $LegendPosition = 'default',


        [string[]] $Colors,
        [string] $Title,
        [ValidateSet('center', 'left', 'right', 'default')][string] $TitleAlignment = 'default',
        [switch] $PatternedColors,
        [switch] $GradientColors,
        [System.Collections.IDictionary] $GridOptions,
        [System.Collections.IDictionary] $Toolbar,
        [System.Collections.IDictionary] $Theme

    )

    $Options = [ordered] @{ }
    New-ChartInternalPie -Options $Options -Names $DataNames -Values $Data -Type $Type


    New-ChartInternalColors -Options $Options -Colors $Colors
    # Default for all charts
    if ($PatternedColors) { New-ChartInternalPattern }
    if ($GradientColors) { New-ChartInternalGradient }
    New-ChartInternalTitle -Options $Options -Title $Title -TitleAlignment $TitleAlignment
    New-ChartInternalSize -Options $Options -Height $Height -Width $Width
    if ($GridOptions) { New-ChartInternalGrid -Options $Options @GridOptions }
    if ($Theme) { New-ChartInternalTheme -Options $Options @Theme }
    if ($Toolbar) { New-ChartInternalToolbar -Options $Options @Toolbar -Show $true }
    New-ApexChart -Options $Options
}

Register-ArgumentCompleter -CommandName New-HTMLChartPie -ParameterName DataLabelsColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLChartPie -ParameterName Colors -ScriptBlock { $Script:RGBColors.Keys }
function New-HTMLChartRadial {
    [CmdletBinding()]
    param(
        [nullable[int]] $Height = 350,
        [nullable[int]] $Width,

        [Array] $DataNames,
        [Array] $Data,
        [string] $Type,
        [ValidateSet('FullCircleTop', 'FullCircleBottom', 'FullCircleBottomLeft', 'FullCircleLeft', 'Speedometer', 'SemiCircleGauge')] $CircleType = 'FullCircleTop',
        [string] $LabelAverage,



        [string[]] $Colors,
        [string] $Title,
        [ValidateSet('center', 'left', 'right', 'default')][string] $TitleAlignment = 'default',
        [switch] $PatternedColors,
        [switch] $GradientColors,
        [System.Collections.IDictionary] $GridOptions,
        [System.Collections.IDictionary] $Toolbar,
        [System.Collections.IDictionary] $Theme
    )

    $Options = [ordered] @{ }

    New-ChartInternalRadial -Options $Options -Names $DataNames -Values $Data -Type $Type
    # This controls how the circle starts / left , right and so on
    New-ChartInternalRadialCircleType -Options $Options -CircleType $CircleType
    # This added label. It's useful if there's more then one data
    New-ChartInternalRadialDataLabels -Options $Options -Label $LabelAverage


    New-ChartInternalColors -Options $Options -Colors $Colors
    # Default for all charts
    if ($PatternedColors) { New-ChartInternalPattern }
    if ($GradientColors) { New-ChartInternalGradient }
    New-ChartInternalTitle -Options $Options -Title $Title -TitleAlignment $TitleAlignment
    New-ChartInternalSize -Options $Options -Height $Height -Width $Width
    if ($GridOptions) { New-ChartInternalGrid -Options $Options @GridOptions }
    if ($Theme) { New-ChartInternalTheme -Options $Options @Theme }
    if ($Toolbar) { New-ChartInternalToolbar -Options $Options @Toolbar -Show $true }
    New-ApexChart -Options $Options
}

Register-ArgumentCompleter -CommandName New-HTMLChartRadial -ParameterName Colors -ScriptBlock { $Script:RGBColors.Keys }
function New-HTMLChartSpark {
    [CmdletBinding()]
    param(
        [nullable[int]] $Height = 350,
        [nullable[int]] $Width,

        [string] $Title,
        [ValidateSet('center', 'left', 'right', 'default')][string] $TitleAlignment = 'default',

        # Data to display in Spark
        [Array] $Data,
        [Array] $DataNames,
        [string] $TitleText,
        [string] $SubTitleText,
        [int] $FontSizeTitle = 24,
        [int] $FontSizeSubtitle = 14,
        [string] $Color,

        [switch] $PatternedColors,
        [switch] $GradientColors,
        [System.Collections.IDictionary] $GridOptions,
        [System.Collections.IDictionary] $Toolbar,
        [System.Collections.IDictionary] $Theme
    )

    $Options = [ordered] @{ }

    New-ChartInternalSpark -Options $Options -Color $Color -Title $TitleText -SubTitle $SubTitleText -FontSizeTitle $FontSizeTitle -FontSizeSubtitle $FontSizeSubtitle -Values $Data


    # Default for all charts
    if ($PatternedColors) { New-ChartInternalPattern }
    if ($GradientColors) { New-ChartInternalGradient }
    New-ChartInternalTitle -Options $Options -Title $Title -TitleAlignment $TitleAlignment
    New-ChartInternalSize -Options $Options -Height $Height -Width $Width
    if ($GridOptions) { New-ChartInternalGrid -Options $Options @GridOptions }
    if ($Theme) { New-ChartInternalTheme -Options $Options @Theme }
    if ($Toolbar) { New-ChartInternalToolbar -Options $Options @Toolbar -Show $true }
    New-ApexChart -Options $Options
}

Register-ArgumentCompleter -CommandName New-HTMLChartSpark -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
$Script:RGBColors = [ordered] @{
    None                   = $null
    AirForceBlue           = 93, 138, 168
    Akaroa                 = 195, 176, 145
    AlbescentWhite         = 227, 218, 201
    AliceBlue              = 240, 248, 255
    Alizarin               = 227, 38, 54
    Allports               = 18, 97, 128
    Almond                 = 239, 222, 205
    AlmondFrost            = 159, 129, 112
    Amaranth               = 229, 43, 80
    Amazon                 = 59, 122, 87
    Amber                  = 255, 191, 0
    Amethyst               = 153, 102, 204
    AmethystSmoke          = 156, 138, 164
    AntiqueWhite           = 250, 235, 215
    Apple                  = 102, 180, 71
    AppleBlossom           = 176, 92, 82
    Apricot                = 251, 206, 177
    Aqua                   = 0, 255, 255
    Aquamarine             = 127, 255, 212
    Armygreen              = 75, 83, 32
    Arsenic                = 59, 68, 75
    Astral                 = 54, 117, 136
    Atlantis               = 164, 198, 57
    Atomic                 = 65, 74, 76
    AtomicTangerine        = 255, 153, 102
    Axolotl                = 99, 119, 91
    Azure                  = 240, 255, 255
    Bahia                  = 176, 191, 26
    BakersChocolate        = 93, 58, 26
    BaliHai                = 124, 152, 171
    BananaMania            = 250, 231, 181
    BattleshipGrey         = 85, 93, 80
    BayOfMany              = 35, 48, 103
    Beige                  = 245, 245, 220
    Bermuda                = 136, 216, 192
    Bilbao                 = 42, 128, 0
    BilobaFlower           = 181, 126, 220
    Bismark                = 83, 104, 114
    Bisque                 = 255, 228, 196
    Bistre                 = 61, 43, 31
    Bittersweet            = 254, 111, 94
    Black                  = 0, 0, 0
    BlackPearl             = 31, 38, 42
    BlackRose              = 85, 31, 47
    BlackRussian           = 23, 24, 43
    BlanchedAlmond         = 255, 235, 205
    BlizzardBlue           = 172, 229, 238
    Blue                   = 0, 0, 255
    BlueDiamond            = 77, 26, 127
    BlueMarguerite         = 115, 102, 189
    BlueSmoke              = 115, 130, 118
    BlueViolet             = 138, 43, 226
    Blush                  = 169, 92, 104
    BokaraGrey             = 22, 17, 13
    Bole                   = 121, 68, 59
    BondiBlue              = 0, 147, 175
    Bordeaux               = 88, 17, 26
    Bossanova              = 86, 60, 92
    Boulder                = 114, 116, 114
    Bouquet                = 183, 132, 167
    Bourbon                = 170, 108, 57
    Brass                  = 181, 166, 66
    BrickRed               = 199, 44, 72
    BrightGreen            = 102, 255, 0
    BrightRed              = 146, 43, 62
    BrightTurquoise        = 8, 232, 222
    BrilliantRose          = 243, 100, 162
    BrinkPink              = 250, 110, 121
    BritishRacingGreen     = 0, 66, 37
    Bronze                 = 205, 127, 50
    Brown                  = 165, 42, 42
    BrownPod               = 57, 24, 2
    BuddhaGold             = 202, 169, 6
    Buff                   = 240, 220, 130
    Burgundy               = 128, 0, 32
    BurlyWood              = 222, 184, 135
    BurntOrange            = 255, 117, 56
    BurntSienna            = 233, 116, 81
    BurntUmber             = 138, 51, 36
    ButteredRum            = 156, 124, 56
    CadetBlue              = 95, 158, 160
    California             = 224, 141, 60
    CamouflageGreen        = 120, 134, 107
    Canary                 = 255, 255, 153
    CanCan                 = 217, 134, 149
    CannonPink             = 145, 78, 117
    CaputMortuum           = 89, 39, 32
    Caramel                = 255, 213, 154
    Cararra                = 237, 230, 214
    Cardinal               = 179, 33, 52
    CardinGreen            = 18, 53, 36
    CareysPink             = 217, 152, 160
    CaribbeanGreen         = 0, 222, 164
    Carmine                = 175, 0, 42
    CarnationPink          = 255, 166, 201
    CarrotOrange           = 242, 142, 28
    Cascade                = 141, 163, 153
    CatskillWhite          = 226, 229, 222
    Cedar                  = 67, 48, 46
    Celadon                = 172, 225, 175
    Celeste                = 207, 207, 196
    Cello                  = 55, 79, 107
    Cement                 = 138, 121, 93
    Cerise                 = 222, 49, 99
    Cerulean               = 0, 123, 167
    CeruleanBlue           = 42, 82, 190
    Chantilly              = 239, 187, 204
    Chardonnay             = 255, 200, 124
    Charlotte              = 167, 216, 222
    Charm                  = 208, 116, 139
    Chartreuse             = 127, 255, 0
    ChartreuseYellow       = 223, 255, 0
    ChelseaCucumber        = 135, 169, 107
    Cherub                 = 246, 214, 222
    Chestnut               = 185, 78, 72
    ChileanFire            = 226, 88, 34
    Chinook                = 150, 200, 162
    Chocolate              = 210, 105, 30
    Christi                = 125, 183, 0
    Christine              = 181, 101, 30
    Cinnabar               = 235, 76, 66
    Citron                 = 159, 169, 31
    Citrus                 = 141, 182, 0
    Claret                 = 95, 25, 51
    ClassicRose            = 251, 204, 231
    ClayCreek              = 145, 129, 81
    Clinker                = 75, 54, 33
    Clover                 = 74, 93, 35
    Cobalt                 = 0, 71, 171
    CocoaBrown             = 44, 22, 8
    Cola                   = 60, 48, 36
    ColumbiaBlue           = 166, 231, 255
    CongoBrown             = 103, 76, 71
    Conifer                = 178, 236, 93
    Copper                 = 218, 138, 103
    CopperRose             = 153, 102, 102
    Coral                  = 255, 127, 80
    CoralRed               = 255, 64, 64
    CoralTree              = 173, 111, 105
    Coriander              = 188, 184, 138
    Corn                   = 251, 236, 93
    CornField              = 250, 240, 190
    Cornflower             = 147, 204, 234
    CornflowerBlue         = 100, 149, 237
    Cornsilk               = 255, 248, 220
    Cosmic                 = 132, 63, 91
    Cosmos                 = 255, 204, 203
    CostaDelSol            = 102, 93, 30
    CottonCandy            = 255, 188, 217
    Crail                  = 164, 90, 82
    Cranberry              = 205, 96, 126
    Cream                  = 255, 255, 204
    CreamCan               = 242, 198, 73
    Crimson                = 220, 20, 60
    Crusta                 = 232, 142, 90
    Cumulus                = 255, 255, 191
    Cupid                  = 246, 173, 198
    CuriousBlue            = 40, 135, 200
    Cyan                   = 0, 255, 255
    Cyprus                 = 6, 78, 64
    DaisyBush              = 85, 53, 146
    Dandelion              = 250, 218, 94
    Danube                 = 96, 130, 182
    DarkBlue               = 0, 0, 139
    DarkBrown              = 101, 67, 33
    DarkCerulean           = 8, 69, 126
    DarkChestnut           = 152, 105, 96
    DarkCoral              = 201, 90, 73
    DarkCyan               = 0, 139, 139
    DarkGoldenrod          = 184, 134, 11
    DarkGray               = 169, 169, 169
    DarkGreen              = 0, 100, 0
    DarkGreenCopper        = 73, 121, 107
    DarkGrey               = 169, 169, 169
    DarkKhaki              = 189, 183, 107
    DarkMagenta            = 139, 0, 139
    DarkOliveGreen         = 85, 107, 47
    DarkOrange             = 255, 140, 0
    DarkOrchid             = 153, 50, 204
    DarkPastelGreen        = 3, 192, 60
    DarkPink               = 222, 93, 131
    DarkPurple             = 150, 61, 127
    DarkRed                = 139, 0, 0
    DarkSalmon             = 233, 150, 122
    DarkSeaGreen           = 143, 188, 143
    DarkSlateBlue          = 72, 61, 139
    DarkSlateGray          = 47, 79, 79
    DarkSlateGrey          = 47, 79, 79
    DarkSpringGreen        = 23, 114, 69
    DarkTangerine          = 255, 170, 29
    DarkTurquoise          = 0, 206, 209
    DarkViolet             = 148, 0, 211
    DarkWood               = 130, 102, 68
    DeepBlush              = 245, 105, 145
    DeepCerise             = 224, 33, 138
    DeepKoamaru            = 51, 51, 102
    DeepLilac              = 153, 85, 187
    DeepMagenta            = 204, 0, 204
    DeepPink               = 255, 20, 147
    DeepSea                = 14, 124, 97
    DeepSkyBlue            = 0, 191, 255
    DeepTeal               = 24, 69, 59
    Denim                  = 36, 107, 206
    DesertSand             = 237, 201, 175
    DimGray                = 105, 105, 105
    DimGrey                = 105, 105, 105
    DodgerBlue             = 30, 144, 255
    Dolly                  = 242, 242, 122
    Downy                  = 95, 201, 191
    DutchWhite             = 239, 223, 187
    EastBay                = 76, 81, 109
    EastSide               = 178, 132, 190
    EchoBlue               = 169, 178, 195
    Ecru                   = 194, 178, 128
    Eggplant               = 162, 0, 109
    EgyptianBlue           = 16, 52, 166
    ElectricBlue           = 125, 249, 255
    ElectricIndigo         = 111, 0, 255
    ElectricLime           = 208, 255, 20
    ElectricPurple         = 191, 0, 255
    Elm                    = 47, 132, 124
    Emerald                = 80, 200, 120
    Eminence               = 108, 48, 130
    Endeavour              = 46, 88, 148
    EnergyYellow           = 245, 224, 80
    Espresso               = 74, 44, 42
    Eucalyptus             = 26, 162, 96
    Falcon                 = 126, 94, 96
    Fallow                 = 204, 153, 102
    FaluRed                = 128, 24, 24
    Feldgrau               = 77, 93, 83
    Feldspar               = 205, 149, 117
    Fern                   = 113, 188, 120
    FernGreen              = 79, 121, 66
    Festival               = 236, 213, 64
    Finn                   = 97, 64, 81
    FireBrick              = 178, 34, 34
    FireBush               = 222, 143, 78
    FireEngineRed          = 211, 33, 45
    Flamingo               = 233, 92, 75
    Flax                   = 238, 220, 130
    FloralWhite            = 255, 250, 240
    ForestGreen            = 34, 139, 34
    Frangipani             = 250, 214, 165
    FreeSpeechAquamarine   = 0, 168, 119
    FreeSpeechRed          = 204, 0, 0
    FrenchLilac            = 230, 168, 215
    FrenchRose             = 232, 83, 149
    FriarGrey              = 135, 134, 129
    Froly                  = 228, 113, 122
    Fuchsia                = 255, 0, 255
    FuchsiaPink            = 255, 119, 255
    Gainsboro              = 220, 220, 220
    Gallery                = 219, 215, 210
    Galliano               = 204, 160, 29
    Gamboge                = 204, 153, 0
    Ghost                  = 196, 195, 208
    GhostWhite             = 248, 248, 255
    Gin                    = 216, 228, 188
    GinFizz                = 247, 231, 206
    Givry                  = 230, 208, 171
    Glacier                = 115, 169, 194
    Gold                   = 255, 215, 0
    GoldDrop               = 213, 108, 43
    GoldenBrown            = 150, 113, 23
    GoldenFizz             = 240, 225, 48
    GoldenGlow             = 248, 222, 126
    GoldenPoppy            = 252, 194, 0
    Goldenrod              = 218, 165, 32
    GoldenSand             = 233, 214, 107
    GoldenYellow           = 253, 238, 0
    GoldTips               = 225, 189, 39
    GordonsGreen           = 37, 53, 41
    Gorse                  = 255, 225, 53
    Gossamer               = 49, 145, 119
    GrannySmithApple       = 168, 228, 160
    Gray                   = 128, 128, 128
    GrayAsparagus          = 70, 89, 69
    Green                  = 0, 128, 0
    GreenLeaf              = 76, 114, 29
    GreenVogue             = 38, 67, 72
    GreenYellow            = 173, 255, 47
    Grey                   = 128, 128, 128
    GreyAsparagus          = 70, 89, 69
    GuardsmanRed           = 157, 41, 51
    GumLeaf                = 178, 190, 181
    Gunmetal               = 42, 52, 57
    Hacienda               = 155, 135, 12
    HalfAndHalf            = 232, 228, 201
    HalfBaked              = 95, 138, 139
    HalfColonialWhite      = 246, 234, 190
    HalfPearlLusta         = 240, 234, 214
    HanPurple              = 63, 0, 255
    Harlequin              = 74, 255, 0
    HarleyDavidsonOrange   = 194, 59, 34
    Heather                = 174, 198, 207
    Heliotrope             = 223, 115, 255
    Hemp                   = 161, 122, 116
    Highball               = 134, 126, 54
    HippiePink             = 171, 75, 82
    Hoki                   = 110, 127, 128
    HollywoodCerise        = 244, 0, 161
    Honeydew               = 240, 255, 240
    Hopbush                = 207, 113, 175
    HorsesNeck             = 108, 84, 30
    HotPink                = 255, 105, 180
    HummingBird            = 201, 255, 229
    HunterGreen            = 53, 94, 59
    Illusion               = 244, 152, 173
    InchWorm               = 202, 224, 13
    IndianRed              = 205, 92, 92
    Indigo                 = 75, 0, 130
    InternationalKleinBlue = 0, 24, 168
    InternationalOrange    = 255, 79, 0
    IrisBlue               = 28, 169, 201
    IrishCoffee            = 102, 66, 40
    IronsideGrey           = 113, 112, 110
    IslamicGreen           = 0, 144, 0
    Ivory                  = 255, 255, 240
    Jacarta                = 61, 50, 93
    JackoBean              = 65, 54, 40
    JacksonsPurple         = 46, 45, 136
    Jade                   = 0, 171, 102
    JapaneseLaurel         = 47, 117, 50
    Jazz                   = 93, 43, 44
    JazzberryJam           = 165, 11, 94
    JellyBean              = 68, 121, 142
    JetStream              = 187, 208, 201
    Jewel                  = 0, 107, 60
    Jon                    = 79, 58, 60
    JordyBlue              = 124, 185, 232
    Jumbo                  = 132, 132, 130
    JungleGreen            = 41, 171, 135
    KaitokeGreen           = 30, 77, 43
    Karry                  = 255, 221, 202
    KellyGreen             = 70, 203, 24
    Keppel                 = 93, 164, 147
    Khaki                  = 240, 230, 140
    Killarney              = 77, 140, 87
    KingfisherDaisy        = 85, 27, 140
    Kobi                   = 230, 143, 172
    LaPalma                = 60, 141, 13
    LaserLemon             = 252, 247, 94
    Laurel                 = 103, 146, 103
    Lavender               = 230, 230, 250
    LavenderBlue           = 204, 204, 255
    LavenderBlush          = 255, 240, 245
    LavenderPink           = 251, 174, 210
    LavenderRose           = 251, 160, 227
    LawnGreen              = 124, 252, 0
    LemonChiffon           = 255, 250, 205
    LightBlue              = 173, 216, 230
    LightCoral             = 240, 128, 128
    LightCyan              = 224, 255, 255
    LightGoldenrodYellow   = 250, 250, 210
    LightGray              = 211, 211, 211
    LightGreen             = 144, 238, 144
    LightGrey              = 211, 211, 211
    LightPink              = 255, 182, 193
    LightSalmon            = 255, 160, 122
    LightSeaGreen          = 32, 178, 170
    LightSkyBlue           = 135, 206, 250
    LightSlateGray         = 119, 136, 153
    LightSlateGrey         = 119, 136, 153
    LightSteelBlue         = 176, 196, 222
    LightYellow            = 255, 255, 224
    Lilac                  = 204, 153, 204
    Lime                   = 0, 255, 0
    LimeGreen              = 50, 205, 50
    Limerick               = 139, 190, 27
    Linen                  = 250, 240, 230
    Lipstick               = 159, 43, 104
    Liver                  = 83, 75, 79
    Lochinvar              = 86, 136, 125
    Lochmara               = 38, 97, 156
    Lola                   = 179, 158, 181
    LondonHue              = 170, 152, 169
    Lotus                  = 124, 72, 72
    LuckyPoint             = 29, 41, 81
    MacaroniAndCheese      = 255, 189, 136
    Madang                 = 193, 249, 162
    Madras                 = 81, 65, 0
    Magenta                = 255, 0, 255
    MagicMint              = 170, 240, 209
    Magnolia               = 248, 244, 255
    Mahogany               = 215, 59, 62
    Maire                  = 27, 24, 17
    Maize                  = 230, 190, 138
    Malachite              = 11, 218, 81
    Malibu                 = 93, 173, 236
    Malta                  = 169, 154, 134
    Manatee                = 140, 146, 172
    Mandalay               = 176, 121, 57
    MandarianOrange        = 146, 39, 36
    Mandy                  = 191, 79, 81
    Manhattan              = 229, 170, 112
    Mantis                 = 125, 194, 66
    Manz                   = 217, 230, 80
    MardiGras              = 48, 25, 52
    Mariner                = 57, 86, 156
    Maroon                 = 128, 0, 0
    Matterhorn             = 85, 85, 85
    Mauve                  = 244, 187, 255
    Mauvelous              = 255, 145, 175
    MauveTaupe             = 143, 89, 115
    MayaBlue               = 119, 181, 254
    McKenzie               = 129, 97, 60
    MediumAquamarine       = 102, 205, 170
    MediumBlue             = 0, 0, 205
    MediumCarmine          = 175, 64, 53
    MediumOrchid           = 186, 85, 211
    MediumPurple           = 147, 112, 219
    MediumRedViolet        = 189, 51, 164
    MediumSeaGreen         = 60, 179, 113
    MediumSlateBlue        = 123, 104, 238
    MediumSpringGreen      = 0, 250, 154
    MediumTurquoise        = 72, 209, 204
    MediumVioletRed        = 199, 21, 133
    MediumWood             = 166, 123, 91
    Melon                  = 253, 188, 180
    Merlot                 = 112, 54, 66
    MetallicGold           = 211, 175, 55
    Meteor                 = 184, 115, 51
    MidnightBlue           = 25, 25, 112
    MidnightExpress        = 0, 20, 64
    Mikado                 = 60, 52, 31
    MilanoRed              = 168, 55, 49
    Ming                   = 54, 116, 125
    MintCream              = 245, 255, 250
    MintGreen              = 152, 255, 152
    Mischka                = 168, 169, 173
    MistyRose              = 255, 228, 225
    Moccasin               = 255, 228, 181
    Mojo                   = 149, 69, 53
    MonaLisa               = 255, 153, 153
    Mongoose               = 179, 139, 109
    Montana                = 53, 56, 57
    MoodyBlue              = 116, 108, 192
    MoonYellow             = 245, 199, 26
    MossGreen              = 173, 223, 173
    MountainMeadow         = 28, 172, 120
    MountainMist           = 161, 157, 148
    MountbattenPink        = 153, 122, 141
    Mulberry               = 211, 65, 157
    Mustard                = 255, 219, 88
    Myrtle                 = 25, 89, 5
    MySin                  = 255, 179, 71
    NavajoWhite            = 255, 222, 173
    Navy                   = 0, 0, 128
    NavyBlue               = 2, 71, 254
    NeonCarrot             = 255, 153, 51
    NeonPink               = 255, 92, 205
    Nepal                  = 145, 163, 176
    Nero                   = 20, 20, 20
    NewMidnightBlue        = 0, 0, 156
    Niagara                = 58, 176, 158
    NightRider             = 59, 47, 47
    Nobel                  = 152, 152, 152
    Norway                 = 169, 186, 157
    Nugget                 = 183, 135, 39
    OceanGreen             = 95, 167, 120
    Ochre                  = 202, 115, 9
    OldCopper              = 111, 78, 55
    OldGold                = 207, 181, 59
    OldLace                = 253, 245, 230
    OldLavender            = 121, 104, 120
    OldRose                = 195, 33, 72
    Olive                  = 128, 128, 0
    OliveDrab              = 107, 142, 35
    OliveGreen             = 181, 179, 92
    Olivetone              = 110, 110, 48
    Olivine                = 154, 185, 115
    Onahau                 = 196, 216, 226
    Opal                   = 168, 195, 188
    Orange                 = 255, 165, 0
    OrangePeel             = 251, 153, 2
    OrangeRed              = 255, 69, 0
    Orchid                 = 218, 112, 214
    OuterSpace             = 45, 56, 58
    OutrageousOrange       = 254, 90, 29
    Oxley                  = 95, 167, 119
    PacificBlue            = 0, 136, 220
    Padua                  = 128, 193, 151
    PalatinatePurple       = 112, 41, 99
    PaleBrown              = 160, 120, 90
    PaleChestnut           = 221, 173, 175
    PaleCornflowerBlue     = 188, 212, 230
    PaleGoldenrod          = 238, 232, 170
    PaleGreen              = 152, 251, 152
    PaleMagenta            = 249, 132, 239
    PalePink               = 250, 218, 221
    PaleSlate              = 201, 192, 187
    PaleTaupe              = 188, 152, 126
    PaleTurquoise          = 175, 238, 238
    PaleVioletRed          = 219, 112, 147
    PalmLeaf               = 53, 66, 48
    Panache                = 233, 255, 219
    PapayaWhip             = 255, 239, 213
    ParisDaisy             = 255, 244, 79
    Parsley                = 48, 96, 48
    PastelGreen            = 119, 221, 119
    PattensBlue            = 219, 233, 244
    Peach                  = 255, 203, 164
    PeachOrange            = 255, 204, 153
    PeachPuff              = 255, 218, 185
    PeachYellow            = 250, 223, 173
    Pear                   = 209, 226, 49
    PearlLusta             = 234, 224, 200
    Pelorous               = 42, 143, 189
    Perano                 = 172, 172, 230
    Periwinkle             = 197, 203, 225
    PersianBlue            = 34, 67, 182
    PersianGreen           = 0, 166, 147
    PersianIndigo          = 51, 0, 102
    PersianPink            = 247, 127, 190
    PersianRed             = 192, 54, 44
    PersianRose            = 233, 54, 167
    Persimmon              = 236, 88, 0
    Peru                   = 205, 133, 63
    Pesto                  = 128, 117, 50
    PictonBlue             = 102, 153, 204
    PigmentGreen           = 0, 173, 67
    PigPink                = 255, 218, 233
    PineGreen              = 1, 121, 111
    PineTree               = 42, 47, 35
    Pink                   = 255, 192, 203
    PinkFlare              = 191, 175, 178
    PinkLace               = 240, 211, 220
    PinkSwan               = 179, 179, 179
    Plum                   = 221, 160, 221
    Pohutukawa             = 102, 12, 33
    PoloBlue               = 119, 158, 203
    Pompadour              = 129, 20, 83
    Portage                = 146, 161, 207
    PotPourri              = 241, 221, 207
    PottersClay            = 132, 86, 60
    PowderBlue             = 176, 224, 230
    Prim                   = 228, 196, 207
    PrussianBlue           = 0, 58, 108
    PsychedelicPurple      = 223, 0, 255
    Puce                   = 204, 136, 153
    Pueblo                 = 108, 46, 31
    PuertoRico             = 67, 179, 174
    Pumpkin                = 255, 99, 28
    Purple                 = 128, 0, 128
    PurpleMountainsMajesty = 150, 123, 182
    PurpleTaupe            = 93, 57, 84
    QuarterSpanishWhite    = 230, 224, 212
    Quartz                 = 220, 208, 255
    Quincy                 = 106, 84, 69
    RacingGreen            = 26, 36, 33
    RadicalRed             = 255, 32, 82
    Rajah                  = 251, 171, 96
    RawUmber               = 123, 63, 0
    RazzleDazzleRose       = 254, 78, 218
    Razzmatazz             = 215, 10, 83
    Red                    = 255, 0, 0
    RedBerry               = 132, 22, 23
    RedDamask              = 203, 109, 81
    RedOxide               = 99, 15, 15
    RedRobin               = 128, 64, 64
    RichBlue               = 84, 90, 167
    Riptide                = 141, 217, 204
    RobinsEggBlue          = 0, 204, 204
    RobRoy                 = 225, 169, 95
    RockSpray              = 171, 56, 31
    RomanCoffee            = 131, 105, 83
    RoseBud                = 246, 164, 148
    RoseBudCherry          = 135, 50, 96
    RoseTaupe              = 144, 93, 93
    RosyBrown              = 188, 143, 143
    Rouge                  = 176, 48, 96
    RoyalBlue              = 65, 105, 225
    RoyalHeath             = 168, 81, 110
    RoyalPurple            = 102, 51, 152
    Ruby                   = 215, 24, 104
    Russet                 = 128, 70, 27
    Rust                   = 192, 64, 0
    RusticRed              = 72, 6, 7
    Saddle                 = 99, 81, 71
    SaddleBrown            = 139, 69, 19
    SafetyOrange           = 255, 102, 0
    Saffron                = 244, 196, 48
    Sage                   = 143, 151, 121
    Sail                   = 161, 202, 241
    Salem                  = 0, 133, 67
    Salmon                 = 250, 128, 114
    SandyBeach             = 253, 213, 177
    SandyBrown             = 244, 164, 96
    Sangria                = 134, 1, 17
    SanguineBrown          = 115, 54, 53
    SanMarino              = 80, 114, 167
    SanteFe                = 175, 110, 77
    Sapphire               = 6, 42, 120
    Saratoga               = 84, 90, 44
    Scampi                 = 102, 102, 153
    Scarlet                = 255, 36, 0
    ScarletGum             = 67, 28, 83
    SchoolBusYellow        = 255, 216, 0
    Schooner               = 139, 134, 128
    ScreaminGreen          = 102, 255, 102
    Scrub                  = 59, 60, 54
    SeaBuckthorn           = 249, 146, 69
    SeaGreen               = 46, 139, 87
    Seagull                = 140, 190, 214
    SealBrown              = 61, 12, 2
    Seance                 = 96, 47, 107
    SeaPink                = 215, 131, 127
    SeaShell               = 255, 245, 238
    Selago                 = 250, 230, 250
    SelectiveYellow        = 242, 180, 0
    SemiSweetChocolate     = 107, 68, 35
    Sepia                  = 150, 90, 62
    Serenade               = 255, 233, 209
    Shadow                 = 133, 109, 77
    Shakespeare            = 114, 160, 193
    Shalimar               = 252, 255, 164
    Shamrock               = 68, 215, 168
    ShamrockGreen          = 0, 153, 102
    SherpaBlue             = 0, 75, 73
    SherwoodGreen          = 27, 77, 62
    Shilo                  = 222, 165, 164
    ShipCove               = 119, 139, 165
    Shocking               = 241, 156, 187
    ShockingPink           = 255, 29, 206
    ShuttleGrey            = 84, 98, 111
    Sidecar                = 238, 224, 177
    Sienna                 = 160, 82, 45
    Silk                   = 190, 164, 147
    Silver                 = 192, 192, 192
    SilverChalice          = 175, 177, 174
    SilverTree             = 102, 201, 146
    SkyBlue                = 135, 206, 235
    SlateBlue              = 106, 90, 205
    SlateGray              = 112, 128, 144
    SlateGrey              = 112, 128, 144
    Smalt                  = 0, 48, 143
    SmaltBlue              = 74, 100, 108
    Snow                   = 255, 250, 250
    SoftAmber              = 209, 190, 168
    Solitude               = 235, 236, 240
    Sorbus                 = 233, 105, 44
    Spectra                = 53, 101, 77
    SpicyMix               = 136, 101, 78
    Spray                  = 126, 212, 230
    SpringBud              = 150, 255, 0
    SpringGreen            = 0, 255, 127
    SpringSun              = 236, 235, 189
    SpunPearl              = 170, 169, 173
    Stack                  = 130, 142, 132
    SteelBlue              = 70, 130, 180
    Stiletto               = 137, 63, 69
    Strikemaster           = 145, 92, 131
    StTropaz               = 50, 82, 123
    Studio                 = 115, 79, 150
    Sulu                   = 201, 220, 135
    SummerSky              = 33, 171, 205
    Sun                    = 237, 135, 45
    Sundance               = 197, 179, 88
    Sunflower              = 228, 208, 10
    Sunglow                = 255, 204, 51
    SunsetOrange           = 253, 82, 64
    SurfieGreen            = 0, 116, 116
    Sushi                  = 111, 153, 64
    SuvaGrey               = 140, 140, 140
    Swamp                  = 35, 43, 43
    SweetCorn              = 253, 219, 109
    SweetPink              = 243, 153, 152
    Tacao                  = 236, 177, 118
    TahitiGold             = 235, 97, 35
    Tan                    = 210, 180, 140
    Tangaroa               = 0, 28, 61
    Tangerine              = 228, 132, 0
    TangerineYellow        = 253, 204, 13
    Tapestry               = 183, 110, 121
    Taupe                  = 72, 60, 50
    TaupeGrey              = 139, 133, 137
    TawnyPort              = 102, 66, 77
    TaxBreak               = 79, 102, 106
    TeaGreen               = 208, 240, 192
    Teak                   = 176, 141, 87
    Teal                   = 0, 128, 128
    TeaRose                = 255, 133, 207
    Temptress              = 60, 20, 33
    Tenne                  = 200, 101, 0
    TerraCotta             = 226, 114, 91
    Thistle                = 216, 191, 216
    TickleMePink           = 245, 111, 161
    Tidal                  = 232, 244, 140
    TitanWhite             = 214, 202, 221
    Toast                  = 165, 113, 100
    Tomato                 = 255, 99, 71
    TorchRed               = 255, 3, 62
    ToryBlue               = 54, 81, 148
    Tradewind              = 110, 174, 161
    TrendyPink             = 133, 96, 136
    TropicalRainForest     = 0, 127, 102
    TrueV                  = 139, 114, 190
    TulipTree              = 229, 183, 59
    Tumbleweed             = 222, 170, 136
    Turbo                  = 255, 195, 36
    TurkishRose            = 152, 119, 123
    Turquoise              = 64, 224, 208
    TurquoiseBlue          = 118, 215, 234
    Tuscany                = 175, 89, 62
    TwilightBlue           = 253, 255, 245
    Twine                  = 186, 135, 89
    TyrianPurple           = 102, 2, 60
    Ultramarine            = 10, 17, 149
    UltraPink              = 255, 111, 255
    Valencia               = 222, 82, 70
    VanCleef               = 84, 61, 55
    VanillaIce             = 229, 204, 201
    VenetianRed            = 209, 0, 28
    Venus                  = 138, 127, 128
    Vermilion              = 251, 79, 20
    VeryLightGrey          = 207, 207, 207
    VidaLoca               = 94, 140, 49
    Viking                 = 71, 171, 204
    Viola                  = 180, 131, 149
    ViolentViolet          = 50, 23, 77
    Violet                 = 238, 130, 238
    VioletRed              = 255, 57, 136
    Viridian               = 64, 130, 109
    VistaBlue              = 159, 226, 191
    VividViolet            = 127, 62, 152
    WaikawaGrey            = 83, 104, 149
    Wasabi                 = 150, 165, 60
    Watercourse            = 0, 106, 78
    Wedgewood              = 67, 107, 149
    WellRead               = 147, 61, 65
    Wewak                  = 255, 152, 153
    Wheat                  = 245, 222, 179
    Whiskey                = 217, 154, 108
    WhiskeySour            = 217, 144, 88
    White                  = 255, 255, 255
    WhiteSmoke             = 245, 245, 245
    WildRice               = 228, 217, 111
    WildSand               = 229, 228, 226
    WildStrawberry         = 252, 65, 154
    WildWatermelon         = 255, 84, 112
    WildWillow             = 172, 191, 96
    Windsor                = 76, 40, 130
    Wisteria               = 191, 148, 228
    Wistful                = 162, 162, 208
    Yellow                 = 255, 255, 0
    YellowGreen            = 154, 205, 50
    YellowOrange           = 255, 174, 66
    YourPink               = 244, 194, 194
}
function Add-CustomFormatForDatetimeSorting {
    <#
    .SYNOPSIS

    .DESCRIPTION
        This function adds code to make the datatable columns sortable with different datetime formats.
        Formatting:
        Day (of Month)
        D       -   1 2 ... 30 31
        Do      -   1st 2nd ... 30th 31st
        DD      -   01 02 ... 30 31

        Month
        M       -   1 2 ... 11 12
        Mo      -   1st 2nd ... 11th 12th
        MM      -   01 02 ... 11 12
        MMM     -   Jan Feb ... Nov Dec
        MMMM    -   January February ... November December

        Year
        YY      -   70 71 ... 29 30
        YYYY    -   1970 1971 ... 2029 2030

        Hour
        H       -   0 1 ... 22 23
        HH      -   00 01 ... 22 23
        h       -   1 2 ... 11 12
        hh      -   01 02 ... 11 12

        Minute
        m       -   0 1 ... 58 59
        mm      -   00 01 ... 58 59

        Second
        s       -   0 1 ... 58 59
        ss      -   00 01 ... 58 59

        More formats
        http://momentjs.com/docs/#/displaying/

    .PARAMETER CustomDateTimeFormat
        Array with strings of custom datetime format.
        The string is build from two parts. Format and locale. Locale is optional.
        format explanation: http://momentjs.com/docs/#/displaying/
        locale explanation: http://momentjs.com/docs/#/i18n/


    .LINK
        format explanation: http://momentjs.com/docs/#/displaying/
        locale explanation: http://momentjs.com/docs/#/i18n/
    .Example
        Add-CustomFormatForDatetimeSorting -CustomDateFormat 'dddd, MMMM Do, YYYY','HH:mm MMM D, YY'
    .Example
        Add-CustomFormatForDatetimeSorting -CustomDateFormat 'DD.MM.YYYY HH:mm:ss'
    #>
    [CmdletBinding()]
    param(
        [array]$DateTimeSortingFormat
    )
    if ($DateTimeSortingFormat) {
        [array]$OutputDateTimeSortingFormat = foreach ($format in $DateTimeSortingFormat) {
            "$.fn.dataTable.moment( '$format' );"
        }
    } else {
        # Default localized format
        $OutputDateTimeSortingFormat = "$.fn.dataTable.moment( 'L' );"
    }
    return $OutputDateTimeSortingFormat
}

function Add-TableContent {
    [CmdletBinding()]
    param(
        [System.Collections.Generic.List[PSCustomObject]] $ContentRows,
        [System.Collections.Generic.List[PSCUstomObject]] $ContentStyle,
        [System.Collections.Generic.List[PSCUstomObject]] $ContentTop,
        [System.Collections.Generic.List[PSCUstomObject]] $ContentFormattingInline,
        [string[]] $HeaderNames,
        [Array] $Table
    )

    # This converts inline conditonal formatting into style. It's intensive because it actually scans whole Table
    # During scan it tries to match things and when it finds a match it prepares it for ContentStyling feature
    if ($ContentFormattingInline.Count -gt 0) {
        [Array] $AddStyles = for ($RowCount = 1; $RowCount -lt $Table.Count; $RowCount++) {
            [string[]] $RowData = $Table[$RowCount] -replace '</td></tr>' -replace '<tr><td>' -split '</td><td>'

            for ($ColumnCount = 0; $ColumnCount -lt $RowData.Count; $ColumnCount++) {
                foreach ($ConditionalFormatting in $ContentFormattingInline) {
                    $ColumnIndexHeader = [array]::indexof($HeaderNames.ToUpper(), $($ConditionalFormatting.Name).ToUpper())
                    if ($ColumnIndexHeader -eq $ColumnCount) {

                        if ($ConditionalFormatting.Type -eq 'number' -or $ConditionalFormatting.Type -eq 'decimal') {
                            [decimal] $returnedValueLeft = 0
                            [bool]$resultLeft = [int]::TryParse($RowData[$ColumnCount], [ref]$returnedValueLeft)

                            [decimal]$returnedValueRight = 0
                            [bool]$resultRight = [int]::TryParse($ConditionalFormatting.Value, [ref]$returnedValueRight)

                            if ($resultLeft -and $resultRight) {
                                $SideLeft = $returnedValueLeft
                                $SideRight = $returnedValueRight
                            } else {
                                $SideLeft = $RowData[$ColumnCount]
                                $SideRight = $ConditionalFormatting.Value
                            }

                        } elseif ($ConditionalFormatting.Type -eq 'int') {
                            # Leaving this in althought only NUMBER is used.
                            [int] $returnedValueLeft = 0
                            [bool]$resultLeft = [int]::TryParse($RowData[$ColumnCount], [ref]$returnedValueLeft)

                            [int]$returnedValueRight = 0
                            [bool]$resultRight = [int]::TryParse($ConditionalFormatting.Value, [ref]$returnedValueRight)

                            if ($resultLeft -and $resultRight) {
                                $SideLeft = $returnedValueLeft
                                $SideRight = $returnedValueRight
                            } else {
                                $SideLeft = $RowData[$ColumnCount]
                                $SideRight = $ConditionalFormatting.Value
                            }
                        } else {
                            $SideLeft = $RowData[$ColumnCount]
                            $SideRight = $ConditionalFormatting.Value
                        }

                        if ($ConditionalFormatting.Operator -eq 'gt') {
                            $Pass = $SideLeft -gt $SideRight
                        } elseif ($ConditionalFormatting.Operator -eq 'lt') {
                            $Pass = $SideLeft -lt $SideRight
                        } elseif ($ConditionalFormatting.Operator -eq 'eq') {
                            $Pass = $SideLeft -eq $SideRight
                        } elseif ($ConditionalFormatting.Operator -eq 'le') {
                            $Pass = $SideLeft -le $SideRight
                        } elseif ($ConditionalFormatting.Operator -eq 'ge') {
                            $Pass = $SideLeft -ge $SideRight
                        } elseif ($ConditionalFormatting.Operator -eq 'ne') {
                            $Pass = $SideLeft -ne $SideRight
                        } elseif ($ConditionalFormatting.Operator -eq 'like') {
                            $Pass = $SideLeft -like $SideRight
                        } elseif ($ConditionalFormatting.Operator -eq 'contains') {
                            $Pass = $SideLeft -contains $SideRight
                        }
                        # This is generally risky, alternative way to do it, so doing above instead
                        # if (Invoke-Expression -Command "`"$($RowData[$ColumnCount])`" -$($ConditionalFormatting.Operator) `"$($ConditionalFormatting.Value)`"") {

                        if ($Pass) {
                            # if ($RowData[$ColumnCount] -eq $ConditionalFormatting.Value) {
                            # If we want to make conditional formatting for for row it requires a bit diff approach
                            if ($ConditionalFormatting.Row) {
                                for ($i = 0; $i -lt $RowData.Count; $i++) {
                                    [PSCustomObject]@{
                                        RowIndex    = $RowCount
                                        ColumnIndex = ($i + 1)
                                        # Since it's 0 based index and we count from 1 we need to add 1
                                        Style       = $ConditionalFormatting.Style
                                    }
                                }
                            } else {
                                [PSCustomObject]@{
                                    RowIndex    = $RowCount
                                    ColumnIndex = ($ColumnIndexHeader + 1)
                                    # Since it's 0 based index and we count from 1 we need to add 1
                                    Style       = $ConditionalFormatting.Style
                                }
                            }
                        }
                    }
                }
            }
        }
        # This makes conditional forwarding a ContentStyle
        foreach ($Style in $AddStyles) {
            $ContentStyle.Add($Style)
        }
    }

    # Prepopulate hashtable with rows
    $TableRows = @{ }
    if ($ContentStyle) {
        for ($RowIndex = 0; $RowIndex -lt $Table.Count; $RowIndex++) {
            $TableRows[$RowIndex] = @{ }
        }
    }

    # Find rows in hashtable and add column to it
    foreach ($Content in $ContentStyle) {
        if ($Content.RowIndex -and $Content.ColumnIndex) {

            # ROWINDEX and COLUMNINDEX - ARRAYS
            # This takes care of Content by Column Nr
            foreach ($ColumnIndex in $Content.ColumnIndex) {

                # Column Index given by user is from 1 to infinity, Column Index is counted from 0
                # We need to address this by doing - 1
                foreach ($RowIndex in $Content.RowIndex) {
                    $TableRows[$RowIndex][$ColumnIndex - 1] = @{
                        Style = $Content.Style
                    }
                    if ($Content.Text) {
                        if ($Content.Used) {
                            $TableRows[$RowIndex][$ColumnIndex - 1]['Text'] = ''
                            $TableRows[$RowIndex][$ColumnIndex - 1]['Remove'] = $true
                        } else {
                            $TableRows[$RowIndex][$ColumnIndex - 1]['Text'] = $Content.Text
                            $TableRows[$RowIndex][$ColumnIndex - 1]['Remove'] = $false
                            $TableRows[$RowIndex][$ColumnIndex - 1]['ColSpan'] = $($Content.ColumnIndex).Count
                            $TableRows[$RowIndex][$ColumnIndex - 1]['RowSpan'] = $($Content.RowIndex).Count
                            $Content.Used = $true
                        }
                    }
                }
            }
        } elseif ($Content.RowIndex -and $Content.Name) {
            # ROWINDEX AND COLUMN NAMES - ARRAYS
            # This takes care of Content by Column Names (Header Names)
            foreach ($ColumnName in $Content.Name) {
                $ColumnIndex = ([array]::indexof($HeaderNames.ToUpper(), $ColumnName.ToUpper()))
                foreach ($RowIndex in $Content.RowIndex) {
                    $TableRows[$RowIndex][$ColumnIndex] = @{
                        Style = $Content.Style
                    }
                    if ($Content.Text) {
                        if ($Content.Used) {
                            $TableRows[$RowIndex][$ColumnIndex]['Text'] = ''
                            $TableRows[$RowIndex][$ColumnIndex]['Remove'] = $true
                        } else {
                            $TableRows[$RowIndex][$ColumnIndex]['Text'] = $Content.Text
                            $TableRows[$RowIndex][$ColumnIndex]['Remove'] = $false
                            $TableRows[$RowIndex][$ColumnIndex]['ColSpan'] = $($Content.ColumnIndex).Count
                            $TableRows[$RowIndex][$ColumnIndex]['RowSpan'] = $($Content.RowIndex).Count
                            $Content.Used = $true
                        }
                    }
                }
            }
        } elseif ($Content.RowIndex -and (-not $Content.ColumnIndex -and -not $Content.Name)) {
            # Just ROW INDEX
            for ($ColumnIndex = 0; $ColumnIndex -lt $HeaderNames.Count; $ColumnIndex++) {
                foreach ($RowIndex in $Content.RowIndex) {
                    $TableRows[$RowIndex][$ColumnIndex] = @{
                        Style = $Content.Style
                    }
                }
            }
        } elseif (-not $Content.RowIndex -and ($Content.ColumnIndex -or $Content.Name)) {
            # JUST COLUMNINDEX or COLUMNNAMES
            for ($RowIndex = 1; $RowIndex -lt $($Table.Count); $RowIndex++) {
                if ($Content.ColumnIndex) {
                    # JUST COLUMN INDEX
                    foreach ($ColumnIndex in $Content.ColumnIndex) {
                        $TableRows[$RowIndex][$ColumnIndex - 1] = @{
                            Style = $Content.Style
                        }
                    }
                } else {
                    # JUST COLUMN NAMES
                    foreach ($ColumnName in $Content.Name) {
                        $ColumnIndex = [array]::indexof($HeaderNames.ToUpper(), $ColumnName.ToUpper())
                        $TableRows[$RowIndex][$ColumnIndex] = @{
                            Style = $Content.Style
                        }
                    }
                }
            }
        }
    }

    # Row 0 = Table Header
    # This builds table from scratch, skipping rows untouched by styling
    [Array] $NewTable = for ($RowCount = 0; $RowCount -lt $Table.Count; $RowCount++) {
        # No conditional formatting we can process just styling since we don't need values
        # We have column index and row index and that's enough
        # In case of conditional formatting it's different as it works on values
        if ($TableRows[$RowCount]) {
            [string[]] $RowData = $Table[$RowCount] -replace '</td></tr>' -replace '<tr><td>' -split '</td><td>'
            New-HTMLTag -Tag 'tr' {
                for ($ColumnCount = 0; $ColumnCount -lt $RowData.Count; $ColumnCount++) {
                    if ($TableRows[$RowCount][$ColumnCount]) {
                        if (-not $TableRows[$RowCount][$ColumnCount]['Remove']) {
                            if ($TableRows[$RowCount][$ColumnCount]['Text']) {
                                New-HTMLTag -Tag 'td' -Value { $TableRows[$RowCount][$ColumnCount]['Text'] } -Attributes @{
                                    style   = $TableRows[$RowCount][$ColumnCount]['Style']
                                    colspan = if ($TableRows[$RowCount][$ColumnCount]['ColSpan'] -gt 1) { $TableRows[$RowCount][$ColumnCount]['ColSpan'] } else { }
                                    rowspan = if ($TableRows[$RowCount][$ColumnCount]['RowSpan'] -gt 1) { $TableRows[$RowCount][$ColumnCount]['RowSpan'] } else { }
                                }

                                # Version 1 - Alternative version to workaround DataTables.NET
                                # New-HTMLTag -Tag 'td' -Value { $TableRows[$RowCount][$ColumnCount]['Text'] } -Attributes @{
                                #    style   = $TableRows[$RowCount][$ColumnCount]['Style']
                                # }

                            } else {
                                New-HTMLTag -Tag 'td' -Value { $RowData[$ColumnCount] } -Attributes @{
                                    style = $TableRows[$RowCount][$ColumnCount]['Style']
                                }
                            }
                        } else {
                            # RowSpan/ColSpan doesn't work in DataTables.net for content.
                            # This means that this functionality is only good for Non-JS.
                            # Normally you would just remove TD/TD and everything shopuld work
                            # And it does work but only for NON-JS solution

                            # Version 1
                            # Alternative Approach - this assumes the text will be zeroed
                            # From visibility side it will look like an empty cells
                            # However content will be stored only in first cell.
                            # requires removal of colspan/rowspan

                            # New-HTMLTag -Tag 'td' -Value { '' } -Attributes @{
                            #    style = $TableRows[$RowCount][$ColumnCount]['Style']
                            # }

                            # Version 2
                            # Below code was suggested as a workaround - it doesn't wrok
                            # New-HTMLTag -Tag 'td' -Value { }  -Attributes @{
                            #     style = "display: none;"
                            # }
                        }
                    } else {
                        New-HTMLTag -Tag 'td' -Value { $RowData[$ColumnCount] }
                    }
                }
            }
        } else {
            $Table[$RowCount]
        }
    }
    $NewTable
}
function Add-TableFiltering {
    [CmdletBinding()]
    param(
        [bool] $Filtering,
        [ValidateSet('Top', 'Bottom', 'Both')][string]$FilteringLocation = 'Bottom',
        [string] $DataTableName
    )
    $Output = @{ }
    if ($Filtering) {
        # https://datatables.net/examples/api/multi_filter.html

        if ($FilteringLocation -eq 'Bottom') {

            $Output.FilteringTopCode = @"
                // Setup - add a text input to each footer cell
                `$('#$DataTableName tfoot th').each(function () {
                    var title = `$(this).text();
                    `$(this).html('<input type="text" placeholder="' + title + '" />');
                });
"@
            $Output.FilteringBottomCode = @"
                // Apply the search for footer cells
                table.columns().every(function () {
                    var that = this;

                    `$('input', this.footer()).on('keyup change', function () {
                        if (that.search() !== this.value) {
                            that.search(this.value).draw();
                        }
                    });
                });
"@

        } elseif ($FilteringLocation -eq 'Both') {

            $Output.FilteringTopCode = @"
                // Setup - add a text input to each header cell
                `$('#$DataTableName thead th').each(function () {
                    var title = `$(this).text();
                    `$(this).html('<input type="text" placeholder="' + title + '" />');
                });
                // Setup - add a text input to each footer cell
                `$('#$DataTableName tfoot th').each(function () {
                    var title = `$(this).text();
                    `$(this).html('<input type="text" placeholder="' + title + '" />');
                });
"@
            $Output.FilteringBottomCode = @"
                // Apply the search for header cells
                table.columns().eq(0).each(function (colIdx) {
                    `$('input', table.column(colIdx).header()).on('keyup change', function () {
                        table
                            .column(colIdx)
                            .search(this.value)
                            .draw();
                    });

                    `$('input', table.column(colIdx).header()).on('click', function (e) {
                        e.stopPropagation();
                    });
                });
                // Apply the search for footer cells
                table.columns().every(function () {
                    var that = this;

                    `$('input', this.footer()).on('keyup change', function () {
                        if (that.search() !== this.value) {
                            that.search(this.value).draw();
                        }
                    });
                });
"@

        } else {
            # top headers
            $Output.FilteringTopCode = @"
                // Setup - add a text input to each header cell
                `$('#$DataTableName thead th').each(function () {
                    var title = `$(this).text();
                    `$(this).html('<input type="text" placeholder="' + title + '" />');
                });
"@

            $Output.FilteringBottomCode = @"
                // Apply the search for header cells
                table.columns().eq(0).each(function (colIdx) {
                    `$('input', table.column(colIdx).header()).on('keyup change', function () {
                        table
                            .column(colIdx)
                            .search(this.value)
                            .draw();
                    });

                    `$('input', table.column(colIdx).header()).on('click', function (e) {
                        e.stopPropagation();
                    });
                });
"@
        }
    } else {
        $Output.FilteringTopCode = $Output.FilteringBottomCode = '' # assign multiple same values trick
    }
    return $Output
}
function Add-TableHeader {
    [CmdletBinding()]
    param(
        [System.Collections.Generic.List[PSCustomObject]] $HeaderRows,
        [System.Collections.Generic.List[PSCUstomObject]] $HeaderStyle,
        [System.Collections.Generic.List[PSCUstomObject]] $HeaderTop,
        [System.Collections.Generic.List[PSCUstomObject]] $HeaderResponsiveOperations,
        [string[]] $HeaderNames
    )
    if ($HeaderRows.Count -eq 0 -and $HeaderStyle.Count -eq 0 -and $HeaderTop.Count -eq 0 -and $HeaderResponsiveOperations.Count -eq 0) {
        return
    }

    # Prepares for styles to merged headers
    [Array] $MergeColumns = foreach ($Row in $HeaderRows) {
        $Index = foreach ($R in $Row.Names) {
            [array]::indexof($HeaderNames.ToUpper(), $R.ToUpper())
        }
        if ($Index -contains -1) {
            Write-Warning -Message "Table Header can't be processed properly. Names on the list to merge were not found in Table Header."
        } else {
            @{
                Index = $Index
                Title = $Row.Title
                Count = $Index.Count
                Style = $Row.Style
                Used  = $false
            }
        }
    }

    $ResponsiveOperations = @{ }
    foreach ($Row in $HeaderResponsiveOperations) {
        foreach ($_ in $Row.Names) {
            $Index = [array]::indexof($HeaderNames.ToUpper(), $_.ToUpper())
            $ResponsiveOperations[$Index] = @{
                Index                = $Index
                ResponsiveOperations = $Row.ResponsiveOperations
                Used                 = $false
            }
        }
    }

    # Prepares for styles to standard header rows
    $Styles = @{ }
    foreach ($Row in $HeaderStyle) {
        foreach ($_ in $Row.Names) {
            $Index = [array]::indexof($HeaderNames.ToUpper(), $_.ToUpper())
            $Styles[$Index] = @{
                Index = $Index
                Title = $Row.Title
                Count = $Index.Count
                Style = $Row.Style
                Used  = $false
            }
        }
    }


    if ($HeaderTop.Count -gt 0) {
        $UsedColumns = 0
        $ColumnsTotal = $HeaderNames.Count
        $TopHeader = New-HTMLTag -Tag 'tr' {
            foreach ($_ in $HeaderTop) {
                if ($_.ColumnCount -eq 0) {
                    $UsedColumns = $ColumnsTotal - $UsedColumns
                    New-HTMLTag -Tag 'th' -Attributes @{ colspan = $UsedColumns; style = ($_.Style) } -Value { $_.Title }
                } else {
                    if ($_.ColumnCount -le $ColumnsTotal) {
                        $UsedColumns = $UsedColumns + $_.ColumnCount
                    } else {
                        $UsedColumns = - ($ColumnsTotal - $_.ColumnCount)
                    }
                    New-HTMLTag -Tag 'th' -Attributes @{ colspan = $_.ColumnCount; style = ($_.Style) } -Value { $_.Title }
                }

            }
        }
    }


    $AddedHeader = @(
        $NewHeader = [System.Collections.Generic.List[string]]::new()
        $TopHeader
        New-HTMLTag -Tag 'tr' {
            for ($i = 0; $i -lt $HeaderNames.Count; $i++) {
                $Found = $false
                foreach ($_ in $MergeColumns) {
                    if ($_.Index -contains $i) {
                        if ($_.Used -eq $false) {
                            New-HTMLTag -Tag 'th' -Attributes @{ colspan = $_.Count; style = ($_.Style); class = $ResponsiveOperations[$i] } -Value { $_.Title }
                            $_.Used = $true
                            $Found = $true
                        } else {
                            $Found = $true
                            # Do Nothing
                        }
                    }
                }
                if (-not $Found) {
                    if ($MergeColumns.Count -eq 0) {
                        # if there are no columns that are supposed to get a Title (merged Title over 2 or more columns) we remove rowspan completly and just apply style
                        # the style will apply, however if Style will be empty it will be removed by New-HTMLTag function
                        New-HTMLTag -Tag 'th' { $HeaderNames[$i] } -Attributes @{ style = $Styles[$i].Style; class = $ResponsiveOperations[$i].ResponsiveOperations }
                    } else {
                        # Since we're adding Title we need to use Rowspan. Rowspan = 2 means spaning row over 2 rows
                        New-HTMLTag -Tag 'th' { $HeaderNames[$i] } -Attributes @{ rowspan = 2; style = $Styles[$i].Style; class = $ResponsiveOperations[$i].ResponsiveOperations }
                    }
                } else {
                    $Head = New-HTMLTag -Tag 'th' { $HeaderNames[$i] } -Attributes @{ style = $Styles[$i].Style; class = $ResponsiveOperations[$i].ResponsiveOperations }
                    $NewHeader.Add($Head)
                }
            }
        }
        if ($NewHeader.Count) {
            New-HTMLTag -Tag 'tr' {
                $NewHeader
            }
        }
    )
    return $AddedHeader
}
function Add-TableRowGrouping {
    [CmdletBinding()]
    param(
        [string] $DataTableName,
        [System.Collections.IDictionary] $Settings,
        [switch] $Top,
        [switch] $Bottom
    )
    if ($Settings.Count -gt 0) {

        if ($Top) {
            $Output = "var collapsedGroups = {};"
        }
        if ($Bottom) {
            $Output = @"
        `$('#$DataTableName tbody').on('click', 'tr.dtrg-start', function () {
            var name = `$(this).data('name');
            collapsedGroups[name] = !collapsedGroups[name];
            table.draw(false);
        });
"@
        }
        $Output
    }
}
function Add-TableState {
    [CmdletBinding()]
    param(
        [bool] $Filtering,
        [bool] $SavedState,
        [string] $DataTableName,
        [ValidateSet('Top', 'Bottom', 'Both')][string]$FilteringLocation = 'Bottom'
    )
    if ($Filtering -and $SavedState) {
        if ($FilteringLocation -eq 'Top') {
            $Output = @"
                // Setup - Looading text input from SavedState
                `$('#$DataTableName').on('stateLoaded.dt', function(e, settings, data) {
                    settings.aoPreSearchCols.forEach(function(col, index) {
                        if (col.sSearch) setTimeout(function() {
                            `$('#$DataTableName thead th:eq('+index+') input').val(col.sSearch)
                        }, 50)
                    })
                });
"@
        } elseif ($FilteringLocation -eq 'Both') {
            $Output = @"
                // Setup - Looading text input from SavedState
                `$('#$DataTableName').on('stateLoaded.dt', function(e, settings, data) {
                    settings.aoPreSearchCols.forEach(function(col, index) {
                        if (col.sSearch) setTimeout(function() {
                            `$('#$DataTableName thead th:eq('+index+') input').val(col.sSearch)
                        }, 50)
                    })
                });
                // Setup - Looading text input from SavedState
                `$('#$DataTableName').on('stateLoaded.dt', function(e, settings, data) {
                    settings.aoPreSearchCols.forEach(function(col, index) {
                        if (col.sSearch) setTimeout(function() {
                            `$('#$DataTableName tfoot th:eq('+index+') input').val(col.sSearch)
                        }, 50)
                    })
                });
"@

        } else {
            $Output = @"
                // Setup - Looading text input from SavedState
                `$('#$DataTableName').on('stateLoaded.dt', function(e, settings, data) {
                    settings.aoPreSearchCols.forEach(function(col, index) {
                        if (col.sSearch) setTimeout(function() {
                            `$('#$DataTableName tfoot th:eq('+index+') input').val(col.sSearch)
                        }, 50)
                    })
                })
"@

        }
    } else {
        $Output = ''
    }
    return $Output
}
function Convert-TableRowGrouping {
    [CmdletBinding()]
    param(
        [string] $Options,
        [int] $RowGroupingColumnID
    )
    if ($RowGroupingColumnID -gt -1) {

        $TextToReplace = @"
        rowGroup: {
            // Uses the 'row group' plugin
            dataSrc: $RowGroupingColumnID,
            startRender: function (rows, group) {
                var collapsed = !!collapsedGroups[group];

                rows.nodes().each(function (r) {
                    r.style.display = collapsed ? 'none' : '';
                });

                var toggleClass = collapsed ? 'fa-plus-square' : 'fa-minus-square';

                // Add group name to <tr>
                return `$('<tr/>')
                    .append('<td colspan="' + rows.columns()[0].length + '">' + '<span class="fa fa-fw ' + toggleClass + ' toggler"/> ' + group + ' (' + rows.count() + ')</td>')
                    .attr('data-name', group)
                    .toggleClass('collapsed', collapsed);
            },
        },
"@
    } else {
        $TextToReplace = ''
    }
    if ($PSEdition -eq 'Desktop') {
        $TextToFind = '"rowGroup":  "",'
    } else {
        $TextToFind = '"rowGroup": "",'
    }

    $Options = $Options -Replace ($TextToFind, $TextToReplace)
    $Options
}
function New-TableConditionalFormatting {
    [CmdletBinding()]
    param(
        [string] $Options,
        [Array] $ConditionalFormatting,
        [string[]] $Header
    )

    if ($ConditionalFormatting.Count -gt 0) {
        # Conditional - changes PowerShellOperator into JS operator
        foreach ($Formatting in $ConditionalFormatting) {
            if ($Formatting.Operator -eq 'gt') {
                $Formatting.Operator = '>'
            } elseif ($Formatting.Operator -eq 'lt') {
                $Formatting.Operator = '<'
            } elseif ($Formatting.Operator -eq 'eq') {
                $Formatting.Operator = '=='
            } elseif ($Formatting.Operator -eq 'le') {
                $Formatting.Operator = '<='
            } elseif ($Formatting.Operator -eq 'ge') {
                $Formatting.Operator = '>='
            } elseif ($Formatting.Operator -eq 'ne') {
                $Formatting.Operator = '!='
            }
            # Operator like/contains are taken care of below
        }
        $Condition = @(
            '"createdRow": function (row, data, dataIndex, column) {'

            foreach ($Condition in $ConditionalFormatting) {
                $ConditionHeaderNr = $Header.ToLower().IndexOf($($Condition.Name.ToLower()))
                $Style = $Condition.Style | ConvertTo-Json
                [string] $StyleDefinition = ".css($Style)"
                if ($null -eq $Condition.Type -or $Condition.Type -eq 'number' -or $Condition.Type -eq 'int' -or $Condition.Type -eq 'decimal') {
                    "if (data[$ConditionHeaderNr] $($Condition.Operator) $($Condition.Value)) {"
                } elseif ($Condition.Type -eq 'string') {
                    switch ($Condition.Operator) {
                        "contains" {
                            "if (data[$($ConditionHeaderNr)].includes('$($Condition.Value)')) {"
                        }
                        "like" {
                            "if (data[$($ConditionHeaderNr)].includes('$($Condition.Value)')) {"
                        }
                        default {
                            "if (data[$ConditionHeaderNr] $($Condition.Operator) '$($Condition.Value)') {"
                        }
                    }
                } elseif ($Condition.Type -eq 'date') {
                    "if (new Date(data[$ConditionHeaderNr]) $($Condition.Operator) new Date('$($Condition.Value)')) {"
                }
                if ($null -ne $Condition.Row -and $Condition.Row -eq $true) {
                    "`$(column)$($StyleDefinition);"
                } else {
                    "`$(column[$ConditionHeaderNr])$($StyleDefinition);"
                }
                "}"
            }

            '}'
        )
        if ($PSEdition -eq 'Desktop') {
            $TextToFind = '"createdRow":  ""'
        } else {
            $TextToFind = '"createdRow": ""'
        }
        $Options = $Options -Replace ($TextToFind, $Condition)
    }
    return $Options
}
function ConvertFrom-Color {
    [alias('Convert-FromColor')]
    [CmdletBinding()]
    param (
        [ValidateScript( {
                if ($($_ -in $Script:RGBColors.Keys -or $_ -match "^#([A-Fa-f0-9]{6})$" -or $_ -eq "") -eq $false) {
                    throw "The Input value is not a valid colorname nor an valid color hex code."
                } else { $true }
            })]
        [alias('Colors')][string[]] $Color,
        [switch] $AsDecimal
    )
    $Colors = foreach ($C in $Color) {
        $Value = $Script:RGBColors."$C"
        if ($C -match "^#([A-Fa-f0-9]{6})$") {
            return $C
        }
        if ($null -eq $Value) {
            return
        }
        $HexValue = Convert-Color -RGB $Value
        Write-Verbose "Convert-FromColor - Color Name: $C Value: $Value HexValue: $HexValue"
        if ($AsDecimal) {
            [Convert]::ToInt64($HexValue, 16)
        } else {
            "#$($HexValue)"
        }
    }
    $Colors
}
Register-ArgumentCompleter -CommandName ConvertFrom-Color -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
function Convert-Image {
    [CmdletBinding()]
    param(
        [string] $Image
    )

    $ImageFile = Get-ImageFile -Image $Image
    if ($ImageFile) {
        Convert-ImageToBinary -ImageFile $ImageFile
    }
}
function Convert-ImagesToBinary {
    [CmdLetBinding()]
    param(
        [string[]] $Content,
        [string] $Search,
        [string] $ReplacePath
    )
    if ($Content -like "*$Search*") {
        if ($PSEdition -eq 'Core') {
            $ImageContent = Get-Content -AsByteStream -LiteralPath $ReplacePath
        } else {
            $ImageContent = Get-Content -LiteralPath $ReplacePath -Encoding Byte
        }
        $Replace = "data:image/$FileType;base64," + [Convert]::ToBase64String($ImageContent)
        $Content = $Content.Replace($Search, $Replace)
    }
    $Content
}
function Convert-ImageToBinary {
    [CmdletBinding()]
    param(
        [System.IO.FileInfo] $ImageFile
    )

    if ($ImageFile.Extension -eq '.jpg') {
        $FileType = 'jpeg'
    } else {
        $FileType = $ImageFile.Extension.Replace('.', '')
    }
    Write-Verbose "Converting $($ImageFile.FullName) to base64 ($FileType)"

    if ($PSEdition -eq 'Core') {
        $ImageContent = Get-Content -AsByteStream -LiteralPath $ImageFile.FullName
    } else {
        $ImageContent = Get-Content -LiteralPath $ImageFile.FullName -Encoding Byte
    }
    $Output = "data:image/$FileType;base64," + [Convert]::ToBase64String(($ImageContent))
    $Output
}
function Convert-StyleContent {
    [CmdLetBinding()]
    param(
        [string[]] $CSS,
        [string] $ImagesPath,
        [string] $SearchPath
    )

    #Get-ObjectType -Object $CSS -VerboseOnly -Verbose

    $ImageFiles = Get-ChildItem -Path (Join-Path $ImagesPath '\*') -Include *.jpg, *.png, *.bmp #-Recurse
    foreach ($Image in $ImageFiles) {
        #$Image.FullName
        #$Image.Name
        $CSS = Convert-ImagesToBinary -Content $CSS -Search "$SearchPath$($Image.Name)" -ReplacePath $Image.FullName
    }
    return $CSS
}

#

#Convert-StyleContent -ImagesPath "$PSScriptRoot\Resources\Images\DataTables" -SearchPath "DataTables-1.10.18/images/"
function Convert-StyleContent1 {
    param(
        [PSCustomObject] $Options
    )
    # Replace PNG / JPG files in Styles
    if ($null -ne $Options.StyleContent) {
        Write-Verbose "Logos: $($Options.Logos.Keys -join ',')"
        foreach ($Logo in $Options.Logos.Keys) {
            $Search = "../images/$Logo.png", "DataTables-1.10.18/images/$Logo.png"
            $Replace = $Options.Logos[$Logo]
            foreach ($S in $Search) {
                Write-Verbose "Logos - replacing $S with binary representation"
                $Options.StyleContent = ($Options.StyleContent).Replace($S, $Replace)
            }
        }
    }    
}
function ConvertTo-CSS {
    [CmdletBinding()]
    param(
        [string] $ID,
        [string] $ClassName,
        [System.Collections.IDictionary] $Attributes,
        [switch] $Group
    )
    $Css = @(
        if ($Group) {
            '<style>'
        }
        if ($ID) {
            "#$ID $ClassName {"
        } else {
            ".$ClassName {"
        }
        foreach ($_ in $Attributes.Keys) {
            if ($null -ne $Attributes[$_]) {
                "$($_): $($Attributes[$_]);"
            }
        }
        '}'
        if ($Group) {
            '</style>'
        }
    ) -join "`n"
    $CSS
}
function ConvertTo-HTMLStyle {
    [CmdletBinding()]
    param(
        [string]$Color,
        [string]$BackGroundColor,
        [int] $FontSize,
        [ValidateSet('normal', 'bold', 'bolder', 'lighter', '100', '200', '300', '400', '500', '600', '700', '800', '900')][string] $FontWeight,
        [ValidateSet('normal', 'italic', 'oblique')][string] $FontStyle,
        [ValidateSet('normal', 'small-caps')][string] $FontVariant,
        [string] $FontFamily,
        [ValidateSet('left', 'center', 'right', 'justify')][string]  $Alignment,
        [ValidateSet('none', 'line-through', 'overline', 'underline')][string]  $TextDecoration,
        [ValidateSet('uppercase', 'lowercase', 'capitalize')][string]  $TextTransform,
        [ValidateSet('rtl')][string] $Direction,
        [switch] $LineBreak
    )
    if ($FontSize -eq 0) {
        $Size = ''
    } else {
        $size = "$($FontSize)px"
    }
    $Style = @{
        'color'            = ConvertFrom-Color -Color $Color
        'background-color' = ConvertFrom-Color -Color $BackGroundColor
        'font-size'        = $Size
        'font-weight'      = $FontWeight
        'font-variant'     = $FontVariant
        'font-family'      = $FontFamily
        'font-style'       = $FontStyle
        'text-align'       = $Alignment


        'text-decoration'  = $TextDecoration
        'text-transform'   = $TextTransform
    }
    # Removes empty, not needed values from hashtable. It's much easier then using if/else to verify for null/empty string
    Remove-EmptyValues -Hashtable $Style
    return $Style
}
function Get-FeaturesInUse {
    <#
    .SYNOPSIS
    Short description

    .DESCRIPTION
    Long description

    .PARAMETER PriorityFeatures
    Define priority features - important for ordering when CSS or JS has to be processed in certain order

    .EXAMPLE
    Get-FeaturesInUse -PriorityFeatures 'Jquery', 'DataTables', 'Tabs', 'Test'

    .NOTES
    General notes
    #>

    [CmdletBinding()]
    param(
        [string[]] $PriorityFeatures
    )
    [Array] $Features = foreach ($Key in $Script:HTMLSchema.Features.Keys) {
        if ($Script:HTMLSchema.Features[$Key]) {
            $Key
        }
    }
    [Array] $TopFeatures = foreach ($Feature in $PriorityFeatures) {
        if ($Features -contains $Feature) {
            $Feature
        }
    }
    [Array] $RemainingFeatures = foreach ($Feature in $Features) {
        if ($TopFeatures -notcontains $Feature) {
            $Feature
        }
    }
    [Array] $AllFeatures = $TopFeatures + $RemainingFeatures
    $AllFeatures
}
Function Get-HTMLLogos {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$LeftLogoName = "Sample",
        [string]$RightLogoName = "Alternate",
        [string]$LeftLogoString,
        [string]$RightLogoString

    )
    $LogoSources = [ordered] @{ }
    $LogoPath = @(
        if ([String]::IsNullOrEmpty($RightLogoString) -eq $false -or [String]::IsNullOrEmpty($LeftLogoString) -eq $false) {
            if ([String]::IsNullOrEmpty($RightLogoString) -eq $false) {
                $LogoSources.Add($RightLogoName, $RightLogoString)
            }
            if ([String]::IsNullOrEmpty($LeftLogoString) -eq $false) {
                $LogoSources.Add($LeftLogoName, $LeftLogoString)
            }
        } else {
            "$PSScriptRoot\Resources\Images\Other"
        }
        "$PSScriptRoot\Resources\Images\DataTables"
    )
    $ImageFiles = Get-ChildItem -Path (Join-Path $LogoPath '\*') -Include *.jpg, *.png, *.bmp -Recurse
    foreach ($ImageFile in $ImageFiles) {
        <#
        if ($ImageFile.Extension -eq '.jpg') {
            $FileType = 'jpeg'
        } else {
            $FileType = $ImageFile.Extension.Replace('.', '')
        }
        Write-Verbose "Converting $($ImageFile.FullName) to base64 ($FileType)"

        if ($PSEdition -eq 'Core') {
            $ImageContent = Get-Content -AsByteStream -LiteralPath $ImageFile.FullName
        } else {
            $ImageContent = Get-Content -LiteralPath $ImageFile.FullName -Encoding Byte
        }
        #>
        $ImageBinary = Convert-ImageToBinary -ImageFile $ImageFile
        #$LogoSources.Add($ImageFile.BaseName, "data:image/$FileType;base64," + [Convert]::ToBase64String(($ImageContent)))
        $LogoSources.Add($ImageFile.BaseName, $ImageBinary)
    }
    $LogoSources
}


#$t = 'C:\Support\GitHub\PSWriteHTML\Private\Get-HTMLLogos.ps1'
#$t -as [System.IO.FileInfo]
function Get-HTMLPartContent {
    param(
        [Array] $Content,
        [string] $Start,
        [string] $End,
        [ValidateSet('Before', 'Between', 'After')] $Type = 'Between'
    )
    $NrStart = $Content.IndexOf($Start)
    $NrEnd = $Content.IndexOf($End)   
    
    #Write-Color $NrStart, $NrEnd, $Type -Color White, Yellow, Blue

    if ($Type -eq 'Between') {
        if ($NrStart -eq -1) {
            # return nothing
            return
        }
        $Content[$NrStart..$NrEnd]
    } 
    if ($Type -eq 'After') {
        if ($NrStart -eq -1) {
            # Returns untouched content
            return $Content
        }
        $Content[($NrEnd + 1)..($Content.Count - 1)]

    }
    if ($Type -eq 'Before') {
        if ($NrStart -eq -1) {
            # return nothing
            return
        }
        $Content[0..$NrStart]
    }
}
function Get-ImageFile {
    [CmdletBinding()]
    param(
        [uri] $Image
    )
    if (-not $Image.IsFile) {
        $Extension = ($Image.OriginalString).Substring(($Image.OriginalString).Length - 4)

        if ($Extension -notin @('.png', '.jpg', 'jpeg', '.svg')) {
            return
        }
        $Extension = $Extension.Replace('.', '')
        $ImageFile = Get-FileName -Extension $Extension -Temporary
        Invoke-WebRequest -Uri $Image -OutFile $ImageFile
        $ImageFile
    } else {

    }
}
function Get-Resources {
    [CmdLetBinding()]
    param(
        [switch] $UseCssLinks,
        [switch] $UseJavaScriptLinks,
        [switch] $NoScript,
        [ValidateSet('Header', 'Footer', 'HeaderAlways', 'FooterAlways')][string] $Location
    )
    DynamicParam {
        # Defines Features Parameter Dynamically
        $Names = $Script:Configuration.Features.Keys
        $ParamAttrib = New-Object System.Management.Automation.ParameterAttribute
        $ParamAttrib.Mandatory = $true
        $ParamAttrib.ParameterSetName = '__AllParameterSets'

        $ReportAttrib = New-Object  System.Collections.ObjectModel.Collection[System.Attribute]
        $ReportAttrib.Add($ParamAttrib)
        $ReportAttrib.Add((New-Object System.Management.Automation.ValidateSetAttribute($Names)))
        $ReportRuntimeParam = New-Object System.Management.Automation.RuntimeDefinedParameter('Features', [string[]], $ReportAttrib)
        $RuntimeParamDic = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $RuntimeParamDic.Add('Features', $ReportRuntimeParam)
        return $RuntimeParamDic
    }
    Process {
        [string[]] $Features = $PSBoundParameters.Features

        foreach ($Feature in $Features) {

            Write-Verbose "Get-Resources - Location: $Location - Feature: $Feature UseCssLinks: $UseCssLinks UseJavaScriptLinks: $UseJavaScriptLinks"
            if ($UseCssLinks) {
                New-HTMLResourceCSS -Link $Script:Configuration.Features.$Feature.$Location.'CssLink' -ResourceComment $Script:Configuration.Features.$Feature.Comment
            } else {
                $CSSOutput = New-HTMLResourceCSS `
                    -FilePath $Script:Configuration.Features.$Feature.$Location.'Css' `
                    -ResourceComment $Script:Configuration.Features.$Feature.Comment `
                    -Replace $Script:Configuration.Features.$Feature.CustomActionsReplace
                Convert-StyleContent -CSS $CSSOutput -ImagesPath "$PSScriptRoot\Resources\Images\DataTables" -SearchPath "../images/"
            }
            if ($UseJavaScriptLinks) {
                New-HTMLResourceJS -Link $Script:Configuration.Features.$Feature.$Location.'JsLink' -ResourceComment $Script:Configuration.Features.$Feature.Comment
            } else {
                New-HTMLResourceJS -FilePath $Script:Configuration.Features.$Feature.$Location.'Js' -ResourceComment $Script:Configuration.Features.$Feature.Comment -ReplaceData $Script:Configuration.Features.$Feature.CustomActionsReplace
            }

            if ($NoScript) {
                [Array] $Output = @(
                    if ($UseCssLinks) {
                        New-HTMLResourceCSS -Link $Script:Configuration.Features.$Feature.$Location.'CssLinkNoScript' -ResourceComment $Script:Configuration.Features.$Feature.Comment
                    } else {
                        $CSSOutput = New-HTMLResourceCSS -FilePath $Script:Configuration.Features.$Feature.$Location.'CssNoScript' -ResourceComment $Script:Configuration.Features.$Feature.Comment -ReplaceData $Script:Configuration.Features.$Feature.CustomActionsReplace
                        Convert-StyleContent -CSS $CSSOutput -ImagesPath "$PSScriptRoot\Resources\Images\DataTables" -SearchPath "../images/"
                    }
                )
                if (($Output.Count -gt 0) -and ($null -ne $Output[0])) {
                    New-HTMLTag -Tag 'noscript' {
                        $Output
                    }
                }
            }
        }
    }
}
function New-DiagramInternalEvent {
    [CmdletBinding()]
    param(
        #[switch] $OnClick,
        [string] $ID,
        #[switch] $FadeSearch,
        [nullable[int]] $ColumnID
    )
    # not ready
    $FadeSearch = $false
    if ($FadeSearch) {
        $Event = @"
        var table = `$('#$ID').DataTable();
        //table.search(params.nodes).draw();
        table.rows(':visible').every(function (rowIdx, tableLoop, rowLoop) {
            var present = true;
            if (params.nodes) {
                present = table.row(rowIdx).data().some(function (v) {
                        return v.match(new RegExp(params.nodes, 'i')) != null;
                    });
            }
            `$(table.row(rowIdx).node()).toggleClass('notMatched', !present);
        });

"@

    } else {
        if ($null -ne $ColumnID) {
            $Event = @"
        var table = `$('#$ID').DataTable();
        if (findValue != '') {
            table.columns($ColumnID).search("^" + findValue + "$", true, false, true).draw();
        } else {
            table.columns($ColumnID).search('').draw();
        }
        if (table.page.info().recordsDisplay == 0) {
            table.columns($ColumnID).search('').draw();
        }
"@
        } else {
            $Event = @"
        var table = `$('#$ID').DataTable();
        if (findValue != '') {
            table.search("^" + findValue + "$", true, false, true).draw();
        } else {
            table.search('').draw();
        }
        if (table.page.info().recordsDisplay == 0) {
            table.search('').draw();
        }
"@
        }
    }
    $Event
}
function New-HTMLAnchor {
    <#
    .SYNOPSIS
    Short description
    
    .DESCRIPTION
    Long description
    
    .PARAMETER Name
    Parameter description
    
    .PARAMETER Id
    Parameter description
    
    .PARAMETER Target
    Parameter description
    
    .PARAMETER Class
    Parameter description
    
    .PARAMETER HrefLink
    Parameter description
    
    .PARAMETER OnClick
    Parameter description
    
    .PARAMETER Style
    Parameter description
    
    .PARAMETER Text
    Parameter description
    
    .EXAMPLE
    New-HTMLAnchor -Target _parent

    New-HTMLAnchor -Id "show_$RandomNumber" -Href '#' -OnClick "show('$RandomNumber');" -Style "color: #ffffff; display:none;" -Text 'Show' 

    Output:
    <a target = "_parent" />
    
    .NOTES
    General notes
    #>
    

    param(
        [alias('AnchorName')][string] $Name,
        [string] $Id,
        [string] $Target, # "_blank|_self|_parent|_top|framename"
        [string] $Class,
        [alias('Url', 'Link', 'UrlLink', 'Href')][string] $HrefLink,
        [string] $OnClick,
        [string] $Style,
        [alias('AnchorText', 'Value')][string] $Text
    )
    $Attributes = [ordered]@{
        'id'      = $Id
        'name'    = $Name
        'class'   = $Class
        'target'  = $Target
        'href'    = $HrefLink
        'onclick' = $OnClick
        'style'   = $Style
    }
    New-HTMLTag -Tag 'a' -Attributes $Attributes {
        $Text
    }
}
function New-HTMLCustomCSS {
    [CmdletBinding()]
    param(
        [System.Collections.IList] $CSS
    )
    "<!-- CSS AUTOGENERATED on DEMAND START -->"
    foreach ($_ in $CSS) {
        if ($_) {
            New-HTMLTag -Tag 'style' -Attributes @{ type = 'text/css' } {
                $_
            } -NewLine
        }
    }
    "<!-- CSS AUTOGENERATED on DEMAND END -->"
}
function New-HTMLResourceCSS {
    [alias('New-ResourceCSS', 'New-CSS')]
    [CmdletBinding()]
    param(
        [alias('ScriptContent')][Parameter(Mandatory = $false, Position = 0)][ValidateNotNull()][ScriptBlock] $Content,
        [string[]] $Link,
        [string] $ResourceComment,
        [string[]] $FilePath,
        [System.Collections.IDictionary] $ReplaceData

    )
    $Output = @(
        "<!-- CSS $ResourceComment START -->"
        foreach ($File in $FilePath) {
            if ($File -ne '') {
                if (Test-Path -LiteralPath $File) {
                    New-HTMLTag -Tag 'style' -Attributes @{ type = 'text/css' } {
                        Write-Verbose "New-HTMLResourceCSS - Reading file from $File"
                        # Replaces stuff based on $Script:Configuration CustomActionReplace Entry
                        $FileContent = Get-Content -LiteralPath $File
                        if ($null -ne $ReplaceData) {
                            foreach ($_ in $ReplaceData.Keys) {
                                $FileContent = $FileContent -replace $_, $ReplaceData[$_]
                            }
                        }
                        $FileContent -replace '@charset "UTF-8";'
                    } -NewLine
                }
            }
        }
        foreach ($L in $Link) {
            if ($L -ne '') {
                Write-Verbose "New-HTMLResourceCSS - Adding link $L"
                New-HTMLTag -Tag 'link' -Attributes @{ rel = "stylesheet"; type = "text/css"; href = $L } -SelfClosing -NewLine
            }
        }
        "<!-- CSS $ResourceComment END -->"
    )
    if ($Output.Count -gt 2) {
        # Outputs only if more than comments
        $Output
    }
}
function New-HTMLResourceJS {
    [alias('New-ResourceJS', 'New-JavaScript')]
    [CmdletBinding()]
    param(
        [alias('ScriptContent')][Parameter(Mandatory = $false, Position = 0)][ValidateNotNull()][ScriptBlock] $Content,
        [string[]] $Link,
        [string] $ResourceComment,
        [string[]] $FilePath,
        [System.Collections.IDictionary] $ReplaceData
    )
    $Output = @(
        "<!-- JS $ResourceComment START -->"
        foreach ($File in $FilePath) {
            if ($File -ne '') {
                if (Test-Path -LiteralPath $File) {
                    New-HTMLTag -Tag 'script' -Attributes @{ type = 'text/javascript' } {
                        # Replaces stuff based on $Script:Configuration CustomActionReplace Entry
                        $FileContent = Get-Content -LiteralPath $File
                        if ($null -ne $ReplaceData) {
                            foreach ($_ in $ReplaceData.Keys) {
                                $FileContent = $FileContent -replace $_, $ReplaceData[$_]
                            }
                        }
                        $FileContent
                    } -NewLine
                } else {
                    return
                }
            }
        }
        foreach ($L in $Link) {
            if ($L -ne '') {
                New-HTMLTag -Tag 'script' -Attributes @{ type = "text/javascript"; src = $L } -NewLine
            } else {
                return
            }
        }
        "<!-- JS $ResourceComment END -->"
    )
    if ($Output.Count -gt 2) {
        # Outputs only if more than comments
        $Output
    }
}
function New-HTMLTabHead {
    [CmdletBinding()]
    Param (
        [Array] $TabsCollection
    )

    if ($Script:HTMLSchema.TabOptions.SlimTabs) {
        $Style = 'display: inline-block;' # makes tabs wrapperr slim/small
    } else {
        $Style = '' # makes it full-width
    }
    <#
    New-HTMLTag -Tag 'div' -Attributes @{ class = 'tabsWrapper' } {
        New-HTMLTag -Tag 'div' -Attributes @{ class = 'tabs' ; style = $Style } {
            New-HTMLTag -Tag 'div' -Attributes @{ class = 'selector' }
            foreach ($Tab in $Script:HTMLSchema.TabsHeaders) {
                $AttributesA = @{
                    'href'    = 'javascript:void(0)'
                    'data-id' = "$($Tab.Id)"
                }
                if ($Tab.Active) {
                    $AttributesA.class = 'active'
                } else {
                    $AttributesA.class = ''
                }
                New-HTMLTag -Tag 'a' -Attributes $AttributesA {
                    New-HTMLTag -Tag 'div' -Attributes @{ class = $($Tab.Icon); style = $($Tab.StyleIcon) }
                    New-HTMLTag -Tag 'span' -Attributes @{ style = $($Tab.StyleText ) } -Value { $Tab.Name }
                }
            }
        }
    }
    #>

    if ($TabsCollection.Count -gt 0) {
        $Tabs = $TabsCollection
    } else {
        $Tabs = $Script:HTMLSchema.TabsHeaders
    }
    New-HTMLTag -Tag 'div' -Attributes @{ class = 'tabsWrapper' } {
        New-HTMLTag -Tag 'div' -Attributes @{ class = 'tabs' ; style = $Style } {
            New-HTMLTag -Tag 'div' -Attributes @{ 'data-tabs' = 'true'; Style = $Script:BorderStyle } {
                foreach ($Tab in $Tabs) {
                    if ($Tab.Active) {
                        $TabActive = 'active'
                    } else {
                        $TabActive = ''
                    }
                    New-HTMLTag -Tag 'div' -Attributes @{ id = $Tab.ID; class = $TabActive; Style = @{'border-radius' = $Script:BorderStyle.'border-radius' } } {
                        New-HTMLTag -Tag 'div' -Attributes @{ class = $($Tab.Icon); style = $($Tab.StyleIcon) }
                        New-HTMLTag -Tag 'span' -Attributes @{ style = $($Tab.StyleText ) } -Value { $Tab.Name }
                    }
                }
            }
        }
    }

}

function New-InternalDiagram {
    [CmdletBinding()]
    param(
        [System.Collections.IList] $Nodes,
        [System.Collections.IList] $Edges,
        [System.Collections.IList] $Events,
        [System.Collections.IDictionary] $Options,
        [string] $Height,
        [string] $Width,
        [string] $BackgroundImage,
        [string] $BackgroundSize = '100% 100%'
    )
    $Script:HTMLSchema.Features.VisNetwork = $true

    $Style = @{ }
    if ($Width -or $Height) {
        $Style['width'] = $Width
        $Style['height'] = $Height
    }
    if ($BackgroundImage) {
        $Style['background'] = "url('$BackgroundImage')"
        $Style['background-size'] = $BackgroundSize
    }

    [string] $ID = "Diagram-" + (Get-RandomStringName -Size 8)
    $Div = New-HTMLTag -Tag 'div' -Attributes @{ id = $ID; class = 'diagram'; style = $Style }

    $ConvertedNodes = $Nodes -join ', '
    $ConvertedEdges = $Edges -join ', '

    if ($Events.Count -gt 0) {
        [Array] $PreparedEvents = @(
            # https://stackoverflow.com/questions/3446170/escape-string-for-use-in-javascript-regex'
            @'
            function escapeRegExp(string) {
                return string.toString().replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
            };
'@
            'network.on("click", function (params) {'
            'params.event = "[original event]";'
            'var findValue = escapeRegExp(params.nodes);'
            foreach ($_ in $Events) {
                New-DiagramInternalEvent -ID $_.ID -ColumnID $_.ColumnID
            }
            '});'
        )
    }


    $Script = New-HTMLTag -Tag 'script' -Value {
        # Convert Dictionary to JSON and return chart within SCRIPT tag
        # Make sure to return with additional empty string

        '// create an array with nodes'
        "var nodes = new vis.DataSet([$ConvertedNodes]); "

        '// create an array with edges'
        "var edges = new vis.DataSet([$ConvertedEdges]); "

        '// create a network'
        "var Container = document.getElementById('$ID'); "
        "var data = { "
        "   nodes: nodes, "
        "   edges: edges"
        " }; "

        if ($Options) {
            $ConvertedOptions = $Options | ConvertTo-Json -Depth 5 | ForEach-Object { [System.Text.RegularExpressions.Regex]::Unescape($_) }
            "var options = $ConvertedOptions; "
        } else {
            "var options = { }; "
        }
        'var network = new vis.Network(Container, data, options); '

        $PreparedEvents
    } -NewLine

    $Div
    $Script:HTMLSchema.Diagrams.Add($Script)
}
$Script:Configuration = [ordered] @{
    Features = [ordered] @{
        Default                 = @{
            Comment      = 'Always Required Default Visual Settings'
            HeaderAlways = @{
                Css = "$PSScriptRoot\Resources\CSS\default.css"
            }
        }
        DefaultHeadings         = @{
            Comment      = 'Always Required Default Headings'
            HeaderAlways = @{
                Css = "$PSScriptRoot\Resources\CSS\headings.css"
            }
        }
        Accordion               = @{
            Comment      = 'Accordion'
            HeaderAlways = @{
                Css = "$PSScriptRoot\Resources\CSS\accordion-1.css"
            }
        }
        CodeBlocks              = @{
            Comment      = 'EnlighterJS CodeBlocks'
            Header       = @{
                CssLink = 'https://evotec.xyz/wp-content/uploads/pswritehtml/enlighterjs30/enlighterjs.min.css'
                Css     = "$PSScriptRoot\Resources\CSS\enlighterjs.min.css"
                JsLink  = 'https://evotec.xyz/wp-content/uploads/pswritehtml/enlighterjs30/enlighterjs.min.js'
                JS      = "$PSScriptRoot\Resources\JS\enlighterjs.min.js"
            }
            Footer       = @{

            }
            HeaderAlways = @{
                Css = "$PSScriptRoot\Resources\CSS\enlighterjs.css"
            }
            FooterAlways = @{
                JS = "$PSScriptRoot\Resources\JS\enlighterjs-footer.js"
            }
        }
        ChartsApex              = @{
            Comment      = 'Apex Charts'
            Header       = @{
                JsLink = 'https://cdn.jsdelivr.net/npm/apexcharts@3.11.1/dist/apexcharts.min.js'
                JS     = "$PSScriptRoot\Resources\JS\apexcharts.min.js"
            }
            HeaderAlways = @{
                #Css = "$PSScriptRoot\Resources\CSS\apexcharts.css"
            }
        }
        ChartsXkcd              = @{
            Header = @{
                JsLink = @(
                    'https://cdn.jsdelivr.net/npm/chart.xkcd@1.1.12/dist/chart.xkcd.min.js'
                )
                Js     = @(
                    "$PSScriptRoot\Resources\JS\chart.xkcd.min.js"
                )
            }
        }
        Jquery                  = @{
            Comment = 'Jquery'
            Header  = @{
                JsLink = 'https://code.jquery.com/jquery-3.4.1.min.js'
                Js     = "$PSScriptRoot\Resources\JS\jquery-3.4.1.min.js"
            }
        }
        DataTablesOld           = @{
            Comment      = 'DataTables'
            HeaderAlways = @{
                Css         = "$PSScriptRoot\Resources\CSS\datatables.css"
                CssNoscript = "$PSScriptRoot\Resources\CSS\datatables.noscript.css"
            }
            Header       = @{
                CssLink = 'https://cdn.datatables.net/v/dt/jq-3.3.1/dt-1.10.18/af-2.3.2/b-1.5.4/b-colvis-1.5.4/b-html5-1.5.4/b-print-1.5.4/cr-1.5.0/fc-3.2.5/fh-3.1.4/kt-2.5.0/r-2.2.2/rg-1.1.0/rr-1.2.4/sc-1.5.0/sl-1.2.6/datatables.min.css'
                Css     = "$PSScriptRoot\Resources\CSS\datatables.min.css"
                JsLink  = @(
                    "https://cdn.datatables.net/v/dt/jq-3.3.1/dt-1.10.18/af-2.3.2/b-1.5.4/b-colvis-1.5.4/b-html5-1.5.4/b-print-1.5.4/cr-1.5.0/fc-3.2.5/fh-3.1.4/kt-2.5.0/r-2.2.2/rg-1.1.0/rr-1.2.4/sc-1.5.0/sl-1.2.6/datatables.min.js"
                    "https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.8.4/moment.min.js"
                    "https://cdn.datatables.net/plug-ins/1.10.19/sorting/datetime-moment.js"
                )
                JS      = @(
                    "$PSScriptRoot\Resources\JS\datatables.min.js"
                    "$PSScriptRoot\Resources\JS\moment.min.js"
                    "$PSScriptRoot\Resources\JS\datetime-moment.js"
                )
            }
        }
        DataTablesSearchFade    = @{
            Comment = 'DataTables SearchFade'
            Header  = @{
                CssLink = 'https://cdn.datatables.net/plug-ins/preview/searchFade/dataTables.searchFade.min.css'
                Css     = "$PSScriptRoot\Resources\CSS\datatablesSearchFade.css"
                JsLink  = "https://cdn.datatables.net/plug-ins/preview/searchFade/dataTables.searchFade.min.js"
                JS      = "$PSScriptRoot\Resources\JS\datatables.SearchFade.min"
            }
        }

        DataTables              = @{
            Comment      = 'DataTables'
            HeaderAlways = @{
                Css         = "$PSScriptRoot\Resources\CSS\datatables.css"
                CssNoscript = "$PSScriptRoot\Resources\CSS\datatables.noscript.css"
            }
            Header       = @{
                CssLink = @(
                    "https://cdn.datatables.net/1.10.20/css/jquery.dataTables.min.css"
                    "https://cdn.datatables.net/autofill/2.3.4/css/autoFill.dataTables.css"
                    "https://cdn.datatables.net/buttons/1.6.1/css/buttons.dataTables.min.css"
                    "https://cdn.datatables.net/colreorder/1.5.2/css/colReorder.dataTables.min.css"
                    "https://cdn.datatables.net/fixedcolumns/3.3.0/css/fixedColumns.dataTables.min.css"
                    "https://cdn.datatables.net/fixedheader/3.1.6/css/fixedHeader.dataTables.min.css"
                    "https://cdn.datatables.net/keytable/2.5.1/css/keyTable.dataTables.min.css"
                    "https://cdn.datatables.net/responsive/2.2.3/css/responsive.dataTables.min.css"
                    "https://cdn.datatables.net/rowgroup/1.1.1/css/rowGroup.dataTables.min.css"
                    "https://cdn.datatables.net/rowreorder/1.2.6/css/rowReorder.dataTables.min.css"
                    "https://cdn.datatables.net/scroller/2.0.1/css/scroller.dataTables.min.css"
                    "https://cdn.datatables.net/select/1.3.1/css/select.dataTables.min.css"
                )
                Css     = @(
                    "$PSScriptRoot\Resources\CSS\jquery.dataTables.min.css"
                    "$PSScriptRoot\Resources\CSS\autoFill.dataTables.css"
                    "$PSScriptRoot\Resources\CSS\buttons.dataTables.min.css"
                    "$PSScriptRoot\Resources\CSS\colReorder.dataTables.min.css"
                    "$PSScriptRoot\Resources\CSS\fixedColumns.dataTables.min.css"
                    "$PSScriptRoot\Resources\CSS\fixedHeader.dataTables.min.css"
                    "$PSScriptRoot\Resources\CSS\keyTable.dataTables.min.css"
                    "$PSScriptRoot\Resources\CSS\responsive.dataTables.min.css"
                    "$PSScriptRoot\Resources\CSS\rowGroup.dataTables.min.css"
                    "$PSScriptRoot\Resources\CSS\rowReorder.dataTables.min.css"
                    "$PSScriptRoot\Resources\CSS\scroller.dataTables.min.css"
                    "$PSScriptRoot\Resources\CSS\select.dataTables.min.css"
                )
                JsLink  = @(
                    #"https://code.jquery.com/jquery-3.3.1.min.js"
                    #"https://cdnjs.cloudflare.com/ajax/libs/jszip/2.5.0/jszip.min.js"
                    #"https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.36/pdfmake.min.js"
                    #"https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.36/vfs_fonts.js"
                    "https://nightly.datatables.net/js/jquery.dataTables.min.js"
                    "https://cdn.datatables.net/autofill/2.3.4/js/dataTables.autoFill.min.js"
                    "https://cdn.datatables.net/buttons/1.6.1/js/dataTables.buttons.min.js"
                    "https://cdn.datatables.net/buttons/1.6.1/js/buttons.colVis.min.js"
                    "https://cdn.datatables.net/buttons/1.6.1/js/buttons.html5.min.js"
                    "https://cdn.datatables.net/buttons/1.6.1/js/buttons.print.min.js"
                    "https://cdn.datatables.net/colreorder/1.5.2/js/dataTables.colReorder.min.js"
                    "https://cdn.datatables.net/fixedcolumns/3.3.0/js/dataTables.fixedColumns.min.js"
                    "https://cdn.datatables.net/fixedheader/3.1.6/js/dataTables.fixedHeader.min.js"
                    "https://cdn.datatables.net/keytable/2.5.1/js/dataTables.keyTable.min.js"
                    "https://cdn.datatables.net/responsive/2.2.3/js/dataTables.responsive.min.js"
                    "https://cdn.datatables.net/rowgroup/1.1.1/js/dataTables.rowGroup.min.js"
                    "https://cdn.datatables.net/rowreorder/1.2.6/js/dataTables.rowReorder.min.js"
                    "https://cdn.datatables.net/scroller/2.0.1/js/dataTables.scroller.min.js"
                    "https://cdn.datatables.net/select/1.3.1/js/dataTables.select.min.js"
                    "https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.24.0/moment.min.js"
                    "https://cdn.datatables.net/plug-ins/1.10.20/sorting/datetime-moment.js"
                )
                JS      = @(
                    "$PSScriptRoot\Resources\JS\jquery.dataTables.min.js"
                    "$PSScriptRoot\Resources\JS\dataTables.autoFill.min.js"
                    "$PSScriptRoot\Resources\JS\dataTables.buttons.min.js"
                    "$PSScriptRoot\Resources\JS\buttons.colVis.min.js"
                    "$PSScriptRoot\Resources\JS\buttons.html5.min.js"
                    "$PSScriptRoot\Resources\JS\buttons.print.min.js"
                    "$PSScriptRoot\Resources\JS\dataTables.colReorder.min.js"
                    "$PSScriptRoot\Resources\JS\dataTables.fixedColumns.min.js"
                    "$PSScriptRoot\Resources\JS\dataTables.fixedHeader.min.js"
                    "$PSScriptRoot\Resources\JS\dataTables.keyTable.min.js"
                    "$PSScriptRoot\Resources\JS\dataTables.responsive.min.js"
                    "$PSScriptRoot\Resources\JS\dataTables.rowGroup.min.js"
                    "$PSScriptRoot\Resources\JS\dataTables.rowReorder.min.js"
                    "$PSScriptRoot\Resources\JS\dataTables.scroller.min.js"
                    "$PSScriptRoot\Resources\JS\dataTables.select.min.js"
                    "$PSScriptRoot\Resources\JS\moment.min.js"
                    "$PSScriptRoot\Resources\JS\datetime-moment.js"
                )
            }
        }
        DataTablesPDF           = @{
            Comment = 'DataTables PDF Features'
            Header  = @{
                JsLink = @(
                    'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.36/pdfmake.min.js'
                    'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.36/vfs_fonts.js'
                )
                Js     = @(
                    "$PSScriptRoot\Resources\JS\pdfmake.min.js"
                    "$PSScriptRoot\Resources\JS\vfs_fonts.js"
                )
            }
        }
        DataTablesExcel         = @{
            Comment = 'DataTables Excel Features'
            Header  = @{
                JsLink = @(
                    'https://cdnjs.cloudflare.com/ajax/libs/jszip/2.5.0/jszip.min.js'
                )
                JS     = @(
                    "$PSScriptRoot\Resources\JS\jszip.min.js"
                )
            }
        }
        DataTablesSimplify      = @{
            Comment      = 'DataTables (not really) - Simplified'
            HeaderAlways = @{
                Css = "$PSScriptRoot\Resources\CSS\datatables.simplify.css"
            }
        }
        D3Mitch                 = @{
            Comment      = 'D3Mitch Feature'
            Header       = @{
                JsLink  = @(
                    #'https://cdn.jsdelivr.net/npm/d3-mitch-tree@1.0.5/lib/d3-mitch-tree.min.js'
                    'https://cdn.jsdelivr.net/gh/deltoss/d3-mitch-tree@1.0.2/dist/js/d3-mitch-tree.min.js'
                )
                CssLink = @(
                    'https://cdn.jsdelivr.net/gh/deltoss/d3-mitch-tree@1.0.2/dist/css/d3-mitch-tree.min.css'
                    'https://cdn.jsdelivr.net/gh/deltoss/d3-mitch-tree@1.0.2/dist/css/d3-mitch-tree-theme-default.min.css'
                )
            }
            HeaderAlways = @{
                Css = "$PSScriptRoot\Resources\CSS\hierarchicalTree.css"
            }
        }
        Fonts                   = @{
            Comment      = 'Default fonts'
            HeaderAlways = @{
                CssLink = 'https://fonts.googleapis.com/css?family=Roboto|Hammersmith+One|Questrial|Oswald'
            }
        }
        FontsAwesome            = @{
            Comment      = 'Default fonts icons'
            HeaderAlways = @{
                CssLink = 'https://use.fontawesome.com/releases/v5.11.2/css/all.css'
            }
            Other        = @{
                Link = 'https://use.fontawesome.com/releases/v5.11.2/svgs/'
            }
        }
        FullCalendar            = @{
            Comment      = 'FullCalendar Basic'
            HeaderAlways = @{
                Css = "$PSScriptRoot\Resources\CSS\calendar.css"
            }
        }
        FullCalendarCore        = @{
            Comment = 'FullCalendar Core'
            Header  = @{
                CssLink = 'https://cdn.jsdelivr.net/npm/@fullcalendar/core@4.3.1/main.min.css'
                Css     = "$PSScriptRoot\Resources\CSS\fullCalendarCore.css"
                JSLink  = 'https://cdn.jsdelivr.net/npm/@fullcalendar/core@4.3.1/main.min.js'
                JS      = "$PSScriptRoot\Resources\JS\fullCalendarCore.js"
            }

        }
        FullCalendarDayGrid     = @{
            Comment = 'FullCalendar DayGrid'
            Header  = @{
                CssLink = 'https://cdn.jsdelivr.net/npm/@fullcalendar/daygrid@4.3.0/main.min.css'
                Css     = "$PSScriptRoot\Resources\CSS\fullCalendarDayGrid.css"
                JSLink  = 'https://cdn.jsdelivr.net/npm/@fullcalendar/daygrid@4.3.0/main.min.js'
                JS      = "$PSScriptRoot\Resources\JS\fullCalendarDayGrid.js"
            }

        }
        FullCalendarInteraction = @{
            Comment = 'FullCalendar Interaction'
            Header  = @{
                JSLink = 'https://cdn.jsdelivr.net/npm/@fullcalendar/interaction@4.3.0/main.min.js'
                JS     = "$PSScriptRoot\Resources\JS\FullCalendarInteraction.js"
            }

        }
        FullCalendarList        = @{
            Comment = 'FullCalendar List'
            Header  = @{
                CssLink = 'https://cdn.jsdelivr.net/npm/@fullcalendar/list@4.3.0/main.min.css'
                Css     = "$PSScriptRoot\Resources\CSS\fullCalendarList.css"
                JSLink  = 'https://cdn.jsdelivr.net/npm/@fullcalendar/list@4.3.0/main.min.js'
                JS      = "$PSScriptRoot\Resources\JS\fullCalendarList.js"
            }

        }
        FullCalendarRRule       = @{
            Comment = 'FullCalendar RRule'
            Header  = @{
                JSLink = 'https://cdn.jsdelivr.net/npm/@fullcalendar/rrule'
                JS     = "$PSScriptRoot\Resources\JS\fullCalendarRRule.js"
                #https://cdn.jsdelivr.net/npm/@fullcalendar/rrule@4.3.0/main.min.js
            }
        }
        FullCalendarTimeGrid    = @{
            Comment = 'FullCalendar TimeGrid'
            Header  = @{
                CssLink = 'https://cdn.jsdelivr.net/npm/@fullcalendar/timegrid@4.3.0/main.min.css'
                Css     = "$PSScriptRoot\Resources\CSS\fullCalendarTimeGrid.css"
                JSLink  = 'https://cdn.jsdelivr.net/npm/@fullcalendar/timegrid@4.3.0/main.min.js'
                JS      = "$PSScriptRoot\Resources\JS\fullCalendarTimeGrid.js"
            }
        }
        FullCalendarTimeLine    = @{
            Comment = 'FullCalendar TimeLine'
            Header  = @{
                CssLink = 'https://cdn.jsdelivr.net/npm/@fullcalendar/timeline@4.3.0/main.min.css'
                Css     = "$PSScriptRoot\Resources\CSS\fullCalendarTimeLine.css"
                JSLink  = 'https://cdn.jsdelivr.net/npm/@fullcalendar/timeline@4.3.0/main.min.js'
                JS      = "$PSScriptRoot\Resources\JS\fullCalendarTimeLine.js"
            }

        }
        HideSection             = @{
            Comment      = 'Hide Section Code'
            HeaderAlways = @{
                JS = "$PSScriptRoot\Resources\JS\HideSection.js"
            }
        }
        FancyTree               = @{
            Header = @{
                JSLink  = @(
                    'https://cdnjs.cloudflare.com/ajax/libs/jquery.fancytree/2.33.0/jquery.fancytree-all-deps.min.js'
                )
                CSSLink = @(
                    'https://cdn.jsdelivr.net/npm/jquery.fancytree@2.33/dist/skin-win8/ui.fancytree.min.css'
                )
            }
        }
        JustGage                = @{
            Comment = 'Just Gage Library'
            Header  = @{
                JSLink = @(
                    'https://cdnjs.cloudflare.com/ajax/libs/raphael/2.3.0/raphael.min.js'
                    'https://cdnjs.cloudflare.com/ajax/libs/justgage/1.3.3/justgage.min.js'
                )
                JS     = @(
                    "$PSScriptRoot\Resources\JS\raphael-min.js"
                    "$PSScriptRoot\Resources\JS\justgage.min.js"
                )
            }
        }
        <#
        JsTree                  = @{
            Header = @{
                JSLink = @(
                    'https://cdnjs.cloudflare.com/ajax/libs/jstree/3.2.1/jstree.min.js'
                )
                CSSLink = @(
                    'https://cdnjs.cloudflare.com/ajax/libs/jstree/3.2.1/themes/default/style.min.css'
                )
                JS = @(
                    "$PSScriptRoot\Resources\JS\stree.min.js"
                )
                CSS = @(
                    "$PSScriptRoot\Resources\CSS\style.min.css"
                )
            }
        }
        #>
        Popper                  = @{
            Comment      = 'Popper and Tooltip for FullCalendar'
            HeaderAlways = @{
                Css = "$PSScriptRoot\Resources\CSS\popper.css"
            }
            Header       = @{
                JSLink = @(
                    'https://unpkg.com/popper.js/dist/umd/popper.min.js'
                    'https://unpkg.com/tooltip.js/dist/umd/tooltip.min.js'
                )
                JS     = @(
                    "$PSScriptRoot\Resources\JS\popper.js"
                    "$PSScriptRoot\Resources\JS\tooltip.js"
                )
            }
        }
        Tabs                    = @{
            Comment              = 'Elastic Tabs'
            HeaderAlways         = @{
                Css = "$PSScriptRoot\Resources\CSS\tabs-elastic.css"
            }
            FooterAlways         = @{
                JS = "$PSScriptRoot\Resources\JS\tabs-elastic.js"
            }
            CustomActionsReplace = @{
                'ColorSelector' = ConvertFrom-Color -Color "DodgerBlue"
                'ColorTarget'   = ConvertFrom-Color -Color "MediumSlateBlue"
            }
        }
        Tabbis                  = @{
            Comment              = 'Elastic Tabbis'
            HeaderAlways         = @{
                Css = "$PSScriptRoot\Resources\CSS\tabbis.css"
            }
            FooterAlways         = @{
                JS = @(
                    "$PSScriptRoot\Resources\JS\tabbis.js"
                    "$PSScriptRoot\Resources\JS\tabbisAdditional.js"
                )
            }
            CustomActionsReplace = @{
                'ColorSelector' = ConvertFrom-Color -Color "DodgerBlue"
                'ColorTarget'   = ConvertFrom-Color -Color "MediumSlateBlue"
            }
        }
        TabbisGradient          = @{
            Comment              = 'Elastic Tabs Gradient'
            FooterAlways         = @{
                Css = "$PSScriptRoot\Resources\CSS\tabs-elastic.gradient.css"
            }
            CustomActionsReplace = @{
                'ColorSelector' = ConvertFrom-Color -Color "DodgerBlue"
                'ColorTarget'   = ConvertFrom-Color -Color "MediumSlateBlue"
            }
        }
        TabbisTransition        = @{
            Comment      = 'Elastic Tabs Transition'
            FooterAlways = @{
                Css = "$PSScriptRoot\Resources\CSS\tabs-elastic.transition.css"
            }
        }
        TimeLine                = @{
            Comment      = 'Timeline Simple'
            HeaderAlways = @{
                Css = "$PSScriptRoot\Resources\CSS\timeline-simple.css"
            }
        }
        Toasts                  = @{
            Comment      = 'Toasts Looking Messages'
            HeaderAlways = @{
                Css = "$PSScriptRoot\Resources\CSS\Toasts.css"
            }
        }
        StatusButtonical        = @{
            Comment      = 'Status Buttonical'
            HeaderAlways = @{
                Css = "$PSScriptRoot\Resources\CSS\status.css"
            }
        }
        TuiGrid                 = @{
            Comment = 'Tui Grid'
            Header  = @{
                Css     = "$PSScriptRoot\Resources\CSS\tuigrid.css"
                CssLink = 'https://cdn.jsdelivr.net/npm/tui-grid@3.5.0/dist/tui-grid.css'
            }
        }
        VisNetwork              = @{
            Comment      = 'VIS Network Dynamic, browser based visualization libraries'
            HeaderAlways = @{
                Css = "$PSScriptRoot\Resources\CSS\vis-network.diagram.css"
            }
            Header       = @{
                JsLink  = 'https://unpkg.com/vis-network@6.4.6/dist/vis-network.min.js'
                Js      = "$PSScriptRoot\Resources\JS\vis-network.min.js"
                Css     = "$PSScriptRoot\Resources\CSS\vis-network.min.css"
                CssLink = 'https://unpkg.com/vis-network@6.4.6/dist/vis-network.min.css'
            }
            FooterAlways = @{
                JS = "$PSScriptRoot\Resources\JS\vis-networkFunctions.js"
            }
        }
    }
}

function Get-ResourcesContentFromWeb {
    param(
        [uri[]] $ResourceLinks,
        [ValidateSet('CSS', 'JS')][string] $Type
    )

    $Output = foreach ($Link in $ResourceLinks) {
        # [uri] $Link = $File
        $Splitted = $Link.OriginalString -split '/'
        $FileName = $Splitted[-1]
        $FilePath = [IO.Path]::Combine('C:\Users\przemyslaw.klys\OneDrive - Evotec\Support\GitHub\PSWriteHTML\Resources', $Type, $FileName)
        $FilePathScriptRoot = -Join ('"', '$PSScriptRoot\Resources\', "$Type", '\', $FileName, '"')
        $FilePathScriptRoot
        #[IO.Path]::Combine('C:\Users\przemyslaw.klys\OneDrive - Evotec\Support\GitHub\PSWriteHTML\Resources\CSS', $FileName)
        Invoke-WebRequest -Uri $Link -OutFile $FilePath
    }
    $Output
}

#Get-ResourcesContentFromWeb -ResourceLinks $($Script:Configuration).Features.DataTables.Header.JsLink -Type 'JS'
#Get-ResourcesContentFromWeb -ResourceLinks $($Script:Configuration).Features.DataTables.Header.CssLink -Type 'CSS'

#Get-ResourcesContentFromWeb -ResourceLinks $($Script:Configuration).Features.VisNetwork.Header.JsLink -Type 'JS'
#Get-ResourcesContentFromWeb -ResourceLinks $($Script:Configuration).Features.VisNetwork.Header.CssLink -Type 'CSS'


#Get-ResourcesContentFromWeb -ResourceLinks $($Script:Configuration).Features.Jquery.Header.JsLink -Type 'JS'
#Get-ResourcesContentFromWeb -ResourceLinks $($Script:Configuration).Features.Jquery.Header.CssLink -Type 'CSS'
#Get-ResourcesContentFromWeb -ResourceLinks $($Script:Configuration).Features.ChartsApex.Header.JsLink -Type 'JS'
#Get-ResourcesContentFromWeb -ResourceLinks $($Script:Configuration).Features.ChartsApex.Header.CssLink -Type 'CSS'
# Another way to access
# https://use.fontawesome.com/releases/v5.11.2/svgs/brands/accessible-icon.svg 

$Global:HTMLIcons = @{
    FontAwesomeBrands  = [ordered] @{
        '500px'                          = 'f26e'
        'accessible-icon'                = 'f368'
        'accusoft'                       = 'f369'
        'acquisitions-incorporated'      = 'f6af'
        'adn'                            = 'f170'
        'adobe'                          = 'f778'
        'adversal'                       = 'f36a'
        'affiliatetheme'                 = 'f36b'
        'airbnb'                         = 'f834'
        'algolia'                        = 'f36c'
        'alipay'                         = 'f642'
        'amazon'                         = 'f270'
        'amazon-pay'                     = 'f42c'
        'amilia'                         = 'f36d'
        'android'                        = 'f17b'
        'angellist'                      = 'f209'
        'angrycreative'                  = 'f36e'
        'angular'                        = 'f420'
        'app-store'                      = 'f36f'
        'app-store-ios'                  = 'f370'
        'apper'                          = 'f371'
        'apple'                          = 'f179'
        'apple-pay'                      = 'f415'
        'artstation'                     = 'f77a'
        'asymmetrik'                     = 'f372'
        'atlassian'                      = 'f77b'
        'audible'                        = 'f373'
        'autoprefixer'                   = 'f41c'
        'avianex'                        = 'f374'
        'aviato'                         = 'f421'
        'aws'                            = 'f375'
        'bandcamp'                       = 'f2d5'
        'battle-net'                     = 'f835'
        'behance'                        = 'f1b4'
        'behance-square'                 = 'f1b5'
        'bimobject'                      = 'f378'
        'bitbucket'                      = 'f171'
        'bitcoin'                        = 'f379'
        'bity'                           = 'f37a'
        'black-tie'                      = 'f27e'
        'blackberry'                     = 'f37b'
        'blogger'                        = 'f37c'
        'blogger-b'                      = 'f37d'
        'bluetooth'                      = 'f293'
        'bluetooth-b'                    = 'f294'
        'bootstrap'                      = 'f836'
        'btc'                            = 'f15a'
        'buffer'                         = 'f837'
        'buromobelexperte'               = 'f37f'
        'buy-n-large'                    = 'f8a6'
        'canadian-maple-leaf'            = 'f785'
        'cc-amazon-pay'                  = 'f42d'
        'cc-amex'                        = 'f1f3'
        'cc-apple-pay'                   = 'f416'
        'cc-diners-club'                 = 'f24c'
        'cc-discover'                    = 'f1f2'
        'cc-jcb'                         = 'f24b'
        'cc-mastercard'                  = 'f1f1'
        'cc-paypal'                      = 'f1f4'
        'cc-stripe'                      = 'f1f5'
        'cc-visa'                        = 'f1f0'
        'centercode'                     = 'f380'
        'centos'                         = 'f789'
        'chrome'                         = 'f268'
        'chromecast'                     = 'f838'
        'cloudscale'                     = 'f383'
        'cloudsmith'                     = 'f384'
        'cloudversify'                   = 'f385'
        'codepen'                        = 'f1cb'
        'codiepie'                       = 'f284'
        'confluence'                     = 'f78d'
        'connectdevelop'                 = 'f20e'
        'contao'                         = 'f26d'
        'cotton-bureau'                  = 'f89e'
        'cpanel'                         = 'f388'
        'creative-commons'               = 'f25e'
        'creative-commons-by'            = 'f4e7'
        'creative-commons-nc'            = 'f4e8'
        'creative-commons-nc-eu'         = 'f4e9'
        'creative-commons-nc-jp'         = 'f4ea'
        'creative-commons-nd'            = 'f4eb'
        'creative-commons-pd'            = 'f4ec'
        'creative-commons-pd-alt'        = 'f4ed'
        'creative-commons-remix'         = 'f4ee'
        'creative-commons-sa'            = 'f4ef'
        'creative-commons-sampling'      = 'f4f0'
        'creative-commons-sampling-plus' = 'f4f1'
        'creative-commons-share'         = 'f4f2'
        'creative-commons-zero'          = 'f4f3'
        'critical-role'                  = 'f6c9'
        'css3'                           = 'f13c'
        'css3-alt'                       = 'f38b'
        'cuttlefish'                     = 'f38c'
        'd-and-d'                        = 'f38d'
        'd-and-d-beyond'                 = 'f6ca'
        'dashcube'                       = 'f210'
        'delicious'                      = 'f1a5'
        'deploydog'                      = 'f38e'
        'deskpro'                        = 'f38f'
        'dev'                            = 'f6cc'
        'deviantart'                     = 'f1bd'
        'dhl'                            = 'f790'
        'diaspora'                       = 'f791'
        'digg'                           = 'f1a6'
        'digital-ocean'                  = 'f391'
        'discord'                        = 'f392'
        'discourse'                      = 'f393'
        'dochub'                         = 'f394'
        'docker'                         = 'f395'
        'draft2digital'                  = 'f396'
        'dribbble'                       = 'f17d'
        'dribbble-square'                = 'f397'
        'dropbox'                        = 'f16b'
        'drupal'                         = 'f1a9'
        'dyalog'                         = 'f399'
        'earlybirds'                     = 'f39a'
        'ebay'                           = 'f4f4'
        'edge'                           = 'f282'
        'elementor'                      = 'f430'
        'ello'                           = 'f5f1'
        'ember'                          = 'f423'
        'empire'                         = 'f1d1'
        'envira'                         = 'f299'
        'erlang'                         = 'f39d'
        'ethereum'                       = 'f42e'
        'etsy'                           = 'f2d7'
        'evernote'                       = 'f839'
        'expeditedssl'                   = 'f23e'
        'facebook'                       = 'f09a'
        'facebook-f'                     = 'f39e'
        'facebook-messenger'             = 'f39f'
        'facebook-square'                = 'f082'
        'fantasy-flight-games'           = 'f6dc'
        'fedex'                          = 'f797'
        'fedora'                         = 'f798'
        'figma'                          = 'f799'
        'firefox'                        = 'f269'
        'first-order'                    = 'f2b0'
        'first-order-alt'                = 'f50a'
        'firstdraft'                     = 'f3a1'
        'flickr'                         = 'f16e'
        'flipboard'                      = 'f44d'
        'fly'                            = 'f417'
        'font-awesome'                   = 'f2b4'
        'font-awesome-alt'               = 'f35c'
        'font-awesome-flag'              = 'f425'
        'fonticons'                      = 'f280'
        'fonticons-fi'                   = 'f3a2'
        'fort-awesome'                   = 'f286'
        'fort-awesome-alt'               = 'f3a3'
        'forumbee'                       = 'f211'
        'foursquare'                     = 'f180'
        'free-code-camp'                 = 'f2c5'
        'freebsd'                        = 'f3a4'
        'fulcrum'                        = 'f50b'
        'galactic-republic'              = 'f50c'
        'galactic-senate'                = 'f50d'
        'get-pocket'                     = 'f265'
        'gg'                             = 'f260'
        'gg-circle'                      = 'f261'
        'git'                            = 'f1d3'
        'git-alt'                        = 'f841'
        'git-square'                     = 'f1d2'
        'github'                         = 'f09b'
        'github-alt'                     = 'f113'
        'github-square'                  = 'f092'
        'gitkraken'                      = 'f3a6'
        'gitlab'                         = 'f296'
        'gitter'                         = 'f426'
        'glide'                          = 'f2a5'
        'glide-g'                        = 'f2a6'
        'gofore'                         = 'f3a7'
        'goodreads'                      = 'f3a8'
        'goodreads-g'                    = 'f3a9'
        'google'                         = 'f1a0'
        'google-drive'                   = 'f3aa'
        'google-play'                    = 'f3ab'
        'google-plus'                    = 'f2b3'
        'google-plus-g'                  = 'f0d5'
        'google-plus-square'             = 'f0d4'
        'google-wallet'                  = 'f1ee'
        'gratipay'                       = 'f184'
        'grav'                           = 'f2d6'
        'gripfire'                       = 'f3ac'
        'grunt'                          = 'f3ad'
        'gulp'                           = 'f3ae'
        'hacker-news'                    = 'f1d4'
        'hacker-news-square'             = 'f3af'
        'hackerrank'                     = 'f5f7'
        'hips'                           = 'f452'
        'hire-a-helper'                  = 'f3b0'
        'hooli'                          = 'f427'
        'hornbill'                       = 'f592'
        'hotjar'                         = 'f3b1'
        'houzz'                          = 'f27c'
        'html5'                          = 'f13b'
        'hubspot'                        = 'f3b2'
        'imdb'                           = 'f2d8'
        'instagram'                      = 'f16d'
        'intercom'                       = 'f7af'
        'internet-explorer'              = 'f26b'
        'invision'                       = 'f7b0'
        'ioxhost'                        = 'f208'
        'itch-io'                        = 'f83a'
        'itunes'                         = 'f3b4'
        'itunes-note'                    = 'f3b5'
        'java'                           = 'f4e4'
        'jedi-order'                     = 'f50e'
        'jenkins'                        = 'f3b6'
        'jira'                           = 'f7b1'
        'joget'                          = 'f3b7'
        'joomla'                         = 'f1aa'
        'js'                             = 'f3b8'
        'js-square'                      = 'f3b9'
        'jsfiddle'                       = 'f1cc'
        'kaggle'                         = 'f5fa'
        'keybase'                        = 'f4f5'
        'keycdn'                         = 'f3ba'
        'kickstarter'                    = 'f3bb'
        'kickstarter-k'                  = 'f3bc'
        'korvue'                         = 'f42f'
        'laravel'                        = 'f3bd'
        'lastfm'                         = 'f202'
        'lastfm-square'                  = 'f203'
        'leanpub'                        = 'f212'
        'less'                           = 'f41d'
        'line'                           = 'f3c0'
        'linkedin'                       = 'f08c'
        'linkedin-in'                    = 'f0e1'
        'linode'                         = 'f2b8'
        'linux'                          = 'f17c'
        'lyft'                           = 'f3c3'
        'magento'                        = 'f3c4'
        'mailchimp'                      = 'f59e'
        'mandalorian'                    = 'f50f'
        'markdown'                       = 'f60f'
        'mastodon'                       = 'f4f6'
        'maxcdn'                         = 'f136'
        'mdb'                            = 'f8ca'
        'medapps'                        = 'f3c6'
        'medium'                         = 'f23a'
        'medium-m'                       = 'f3c7'
        'medrt'                          = 'f3c8'
        'meetup'                         = 'f2e0'
        'megaport'                       = 'f5a3'
        'mendeley'                       = 'f7b3'
        'microsoft'                      = 'f3ca'
        'mix'                            = 'f3cb'
        'mixcloud'                       = 'f289'
        'mizuni'                         = 'f3cc'
        'modx'                           = 'f285'
        'monero'                         = 'f3d0'
        'napster'                        = 'f3d2'
        'neos'                           = 'f612'
        'nimblr'                         = 'f5a8'
        'node'                           = 'f419'
        'node-js'                        = 'f3d3'
        'npm'                            = 'f3d4'
        'ns8'                            = 'f3d5'
        'nutritionix'                    = 'f3d6'
        'odnoklassniki'                  = 'f263'
        'odnoklassniki-square'           = 'f264'
        'old-republic'                   = 'f510'
        'opencart'                       = 'f23d'
        'openid'                         = 'f19b'
        'opera'                          = 'f26a'
        'optin-monster'                  = 'f23c'
        'orcid'                          = 'f8d2'
        'osi'                            = 'f41a'
        'page4'                          = 'f3d7'
        'pagelines'                      = 'f18c'
        'palfed'                         = 'f3d8'
        'patreon'                        = 'f3d9'
        'paypal'                         = 'f1ed'
        'penny-arcade'                   = 'f704'
        'periscope'                      = 'f3da'
        'phabricator'                    = 'f3db'
        'phoenix-framework'              = 'f3dc'
        'phoenix-squadron'               = 'f511'
        'php'                            = 'f457'
        'pied-piper'                     = 'f2ae'
        'pied-piper-alt'                 = 'f1a8'
        'pied-piper-hat'                 = 'f4e5'
        'pied-piper-pp'                  = 'f1a7'
        'pinterest'                      = 'f0d2'
        'pinterest-p'                    = 'f231'
        'pinterest-square'               = 'f0d3'
        'playstation'                    = 'f3df'
        'product-hunt'                   = 'f288'
        'pushed'                         = 'f3e1'
        'python'                         = 'f3e2'
        'qq'                             = 'f1d6'
        'quinscape'                      = 'f459'
        'quora'                          = 'f2c4'
        'r-project'                      = 'f4f7'
        'raspberry-pi'                   = 'f7bb'
        'ravelry'                        = 'f2d9'
        'react'                          = 'f41b'
        'reacteurope'                    = 'f75d'
        'readme'                         = 'f4d5'
        'rebel'                          = 'f1d0'
        'red-river'                      = 'f3e3'
        'reddit'                         = 'f1a1'
        'reddit-alien'                   = 'f281'
        'reddit-square'                  = 'f1a2'
        'redhat'                         = 'f7bc'
        'renren'                         = 'f18b'
        'replyd'                         = 'f3e6'
        'researchgate'                   = 'f4f8'
        'resolving'                      = 'f3e7'
        'rev'                            = 'f5b2'
        'rocketchat'                     = 'f3e8'
        'rockrms'                        = 'f3e9'
        'safari'                         = 'f267'
        'salesforce'                     = 'f83b'
        'sass'                           = 'f41e'
        'schlix'                         = 'f3ea'
        'scribd'                         = 'f28a'
        'searchengin'                    = 'f3eb'
        'sellcast'                       = 'f2da'
        'sellsy'                         = 'f213'
        'servicestack'                   = 'f3ec'
        'shirtsinbulk'                   = 'f214'
        'shopware'                       = 'f5b5'
        'simplybuilt'                    = 'f215'
        'sistrix'                        = 'f3ee'
        'sith'                           = 'f512'
        'sketch'                         = 'f7c6'
        'skyatlas'                       = 'f216'
        'skype'                          = 'f17e'
        'slack'                          = 'f198'
        'slack-hash'                     = 'f3ef'
        'slideshare'                     = 'f1e7'
        'snapchat'                       = 'f2ab'
        'snapchat-ghost'                 = 'f2ac'
        'snapchat-square'                = 'f2ad'
        'soundcloud'                     = 'f1be'
        'sourcetree'                     = 'f7d3'
        'speakap'                        = 'f3f3'
        'speaker-deck'                   = 'f83c'
        'spotify'                        = 'f1bc'
        'squarespace'                    = 'f5be'
        'stack-exchange'                 = 'f18d'
        'stack-overflow'                 = 'f16c'
        'stackpath'                      = 'f842'
        'staylinked'                     = 'f3f5'
        'steam'                          = 'f1b6'
        'steam-square'                   = 'f1b7'
        'steam-symbol'                   = 'f3f6'
        'sticker-mule'                   = 'f3f7'
        'strava'                         = 'f428'
        'stripe'                         = 'f429'
        'stripe-s'                       = 'f42a'
        'studiovinari'                   = 'f3f8'
        'stumbleupon'                    = 'f1a4'
        'stumbleupon-circle'             = 'f1a3'
        'superpowers'                    = 'f2dd'
        'supple'                         = 'f3f9'
        'suse'                           = 'f7d6'
        'swift'                          = 'f8e1'
        'symfony'                        = 'f83d'
        'teamspeak'                      = 'f4f9'
        'telegram'                       = 'f2c6'
        'telegram-plane'                 = 'f3fe'
        'tencent-weibo'                  = 'f1d5'
        'the-red-yeti'                   = 'f69d'
        'themeco'                        = 'f5c6'
        'themeisle'                      = 'f2b2'
        'think-peaks'                    = 'f731'
        'trade-federation'               = 'f513'
        'trello'                         = 'f181'
        'tripadvisor'                    = 'f262'
        'tumblr'                         = 'f173'
        'tumblr-square'                  = 'f174'
        'twitch'                         = 'f1e8'
        'twitter'                        = 'f099'
        'twitter-square'                 = 'f081'
        'typo3'                          = 'f42b'
        'uber'                           = 'f402'
        'ubuntu'                         = 'f7df'
        'uikit'                          = 'f403'
        'umbraco'                        = 'f8e8'
        'uniregistry'                    = 'f404'
        'untappd'                        = 'f405'
        'ups'                            = 'f7e0'
        'usb'                            = 'f287'
        'usps'                           = 'f7e1'
        'ussunnah'                       = 'f407'
        'vaadin'                         = 'f408'
        'viacoin'                        = 'f237'
        'viadeo'                         = 'f2a9'
        'viadeo-square'                  = 'f2aa'
        'viber'                          = 'f409'
        'vimeo'                          = 'f40a'
        'vimeo-square'                   = 'f194'
        'vimeo-v'                        = 'f27d'
        'vine'                           = 'f1ca'
        'vk'                             = 'f189'
        'vnv'                            = 'f40b'
        'vuejs'                          = 'f41f'
        'waze'                           = 'f83f'
        'weebly'                         = 'f5cc'
        'weibo'                          = 'f18a'
        'weixin'                         = 'f1d7'
        'whatsapp'                       = 'f232'
        'whatsapp-square'                = 'f40c'
        'whmcs'                          = 'f40d'
        'wikipedia-w'                    = 'f266'
        'windows'                        = 'f17a'
        'wix'                            = 'f5cf'
        'wizards-of-the-coast'           = 'f730'
        'wolf-pack-battalion'            = 'f514'
        'wordpress'                      = 'f19a'
        'wordpress-simple'               = 'f411'
        'wpbeginner'                     = 'f297'
        'wpexplorer'                     = 'f2de'
        'wpforms'                        = 'f298'
        'wpressr'                        = 'f3e4'
        'xbox'                           = 'f412'
        'xing'                           = 'f168'
        'xing-square'                    = 'f169'
        'y-combinator'                   = 'f23b'
        'yahoo'                          = 'f19e'
        'yammer'                         = 'f840'
        'yandex'                         = 'f413'
        'yandex-international'           = 'f414'
        'yarn'                           = 'f7e3'
        'yelp'                           = 'f1e9'
        'yoast'                          = 'f2b1'
        'youtube'                        = 'f167'
        'youtube-square'                 = 'f431'
        'zhihu'                          = 'f63f'
    }
    FontAwesomeRegular = [ordered] @{
        'address-book'           = 'f2b9'
        'address-card'           = 'f2bb'
        'angry'                  = 'f556'
        'arrow-alt-circle-down'  = 'f358'
        'arrow-alt-circle-left'  = 'f359'
        'arrow-alt-circle-right' = 'f35a'
        'arrow-alt-circle-up'    = 'f35b'
        'bell'                   = 'f0f3'
        'bell-slash'             = 'f1f6'
        'bookmark'               = 'f02e'
        'building'               = 'f1ad'
        'calendar'               = 'f133'
        'calendar-alt'           = 'f073'
        'calendar-check'         = 'f274'
        'calendar-minus'         = 'f272'
        'calendar-plus'          = 'f271'
        'calendar-times'         = 'f273'
        'caret-square-down'      = 'f150'
        'caret-square-left'      = 'f191'
        'caret-square-right'     = 'f152'
        'caret-square-up'        = 'f151'
        'chart-bar'              = 'f080'
        'check-circle'           = 'f058'
        'check-square'           = 'f14a'
        'circle'                 = 'f111'
        'clipboard'              = 'f328'
        'clock'                  = 'f017'
        'clone'                  = 'f24d'
        'closed-captioning'      = 'f20a'
        'comment'                = 'f075'
        'comment-alt'            = 'f27a'
        'comment-dots'           = 'f4ad'
        'comments'               = 'f086'
        'compass'                = 'f14e'
        'copy'                   = 'f0c5'
        'copyright'              = 'f1f9'
        'credit-card'            = 'f09d'
        'dizzy'                  = 'f567'
        'dot-circle'             = 'f192'
        'edit'                   = 'f044'
        'envelope'               = 'f0e0'
        'envelope-open'          = 'f2b6'
        'eye'                    = 'f06e'
        'eye-slash'              = 'f070'
        'file'                   = 'f15b'
        'file-alt'               = 'f15c'
        'file-archive'           = 'f1c6'
        'file-audio'             = 'f1c7'
        'file-code'              = 'f1c9'
        'file-excel'             = 'f1c3'
        'file-image'             = 'f1c5'
        'file-pdf'               = 'f1c1'
        'file-powerpoint'        = 'f1c4'
        'file-video'             = 'f1c8'
        'file-word'              = 'f1c2'
        'flag'                   = 'f024'
        'flushed'                = 'f579'
        'folder'                 = 'f07b'
        'folder-open'            = 'f07c'
        'frown'                  = 'f119'
        'frown-open'             = 'f57a'
        'futbol'                 = 'f1e3'
        'gem'                    = 'f3a5'
        'grimace'                = 'f57f'
        'grin'                   = 'f580'
        'grin-alt'               = 'f581'
        'grin-beam'              = 'f582'
        'grin-beam-sweat'        = 'f583'
        'grin-hearts'            = 'f584'
        'grin-squint'            = 'f585'
        'grin-squint-tears'      = 'f586'
        'grin-stars'             = 'f587'
        'grin-tears'             = 'f588'
        'grin-tongue'            = 'f589'
        'grin-tongue-squint'     = 'f58a'
        'grin-tongue-wink'       = 'f58b'
        'grin-wink'              = 'f58c'
        'hand-lizard'            = 'f258'
        'hand-paper'             = 'f256'
        'hand-peace'             = 'f25b'
        'hand-point-down'        = 'f0a7'
        'hand-point-left'        = 'f0a5'
        'hand-point-right'       = 'f0a4'
        'hand-point-up'          = 'f0a6'
        'hand-pointer'           = 'f25a'
        'hand-rock'              = 'f255'
        'hand-scissors'          = 'f257'
        'hand-spock'             = 'f259'
        'handshake'              = 'f2b5'
        'hdd'                    = 'f0a0'
        'heart'                  = 'f004'
        'hospital'               = 'f0f8'
        'hourglass'              = 'f254'
        'id-badge'               = 'f2c1'
        'id-card'                = 'f2c2'
        'image'                  = 'f03e'
        'images'                 = 'f302'
        'keyboard'               = 'f11c'
        'kiss'                   = 'f596'
        'kiss-beam'              = 'f597'
        'kiss-wink-heart'        = 'f598'
        'laugh'                  = 'f599'
        'laugh-beam'             = 'f59a'
        'laugh-squint'           = 'f59b'
        'laugh-wink'             = 'f59c'
        'lemon'                  = 'f094'
        'life-ring'              = 'f1cd'
        'lightbulb'              = 'f0eb'
        'list-alt'               = 'f022'
        'map'                    = 'f279'
        'meh'                    = 'f11a'
        'meh-blank'              = 'f5a4'
        'meh-rolling-eyes'       = 'f5a5'
        'minus-square'           = 'f146'
        'money-bill-alt'         = 'f3d1'
        'moon'                   = 'f186'
        'newspaper'              = 'f1ea'
        'object-group'           = 'f247'
        'object-ungroup'         = 'f248'
        'paper-plane'            = 'f1d8'
        'pause-circle'           = 'f28b'
        'play-circle'            = 'f144'
        'plus-square'            = 'f0fe'
        'question-circle'        = 'f059'
        'registered'             = 'f25d'
        'sad-cry'                = 'f5b3'
        'sad-tear'               = 'f5b4'
        'save'                   = 'f0c7'
        'share-square'           = 'f14d'
        'smile'                  = 'f118'
        'smile-beam'             = 'f5b8'
        'smile-wink'             = 'f4da'
        'snowflake'              = 'f2dc'
        'square'                 = 'f0c8'
        'star'                   = 'f005'
        'star-half'              = 'f089'
        'sticky-note'            = 'f249'
        'stop-circle'            = 'f28d'
        'sun'                    = 'f185'
        'surprise'               = 'f5c2'
        'thumbs-down'            = 'f165'
        'thumbs-up'              = 'f164'
        'times-circle'           = 'f057'
        'tired'                  = 'f5c8'
        'trash-alt'              = 'f2ed'
        'user'                   = 'f007'
        'user-circle'            = 'f2bd'
        'window-close'           = 'f410'
        'window-maximize'        = 'f2d0'
        'window-minimize'        = 'f2d1'
        'window-restore'         = 'f2d2'
    }
    FontAwesomeSolid   = [ordered] @{
        'ad'                                  = 'f641'
        'address-book'                        = 'f2b9'
        'address-card'                        = 'f2bb'
        'adjust'                              = 'f042'
        'air-freshener'                       = 'f5d0'
        'align-center'                        = 'f037'
        'align-justify'                       = 'f039'
        'align-left'                          = 'f036'
        'align-right'                         = 'f038'
        'allergies'                           = 'f461'
        'ambulance'                           = 'f0f9'
        'american-sign-language-interpreting' = 'f2a3'
        'anchor'                              = 'f13d'
        'angle-double-down'                   = 'f103'
        'angle-double-left'                   = 'f100'
        'angle-double-right'                  = 'f101'
        'angle-double-up'                     = 'f102'
        'angle-down'                          = 'f107'
        'angle-left'                          = 'f104'
        'angle-right'                         = 'f105'
        'angle-up'                            = 'f106'
        'angry'                               = 'f556'
        'ankh'                                = 'f644'
        'apple-alt'                           = 'f5d1'
        'archive'                             = 'f187'
        'archway'                             = 'f557'
        'arrow-alt-circle-down'               = 'f358'
        'arrow-alt-circle-left'               = 'f359'
        'arrow-alt-circle-right'              = 'f35a'
        'arrow-alt-circle-up'                 = 'f35b'
        'arrow-circle-down'                   = 'f0ab'
        'arrow-circle-left'                   = 'f0a8'
        'arrow-circle-right'                  = 'f0a9'
        'arrow-circle-up'                     = 'f0aa'
        'arrow-down'                          = 'f063'
        'arrow-left'                          = 'f060'
        'arrow-right'                         = 'f061'
        'arrow-up'                            = 'f062'
        'arrows-alt'                          = 'f0b2'
        'arrows-alt-h'                        = 'f337'
        'arrows-alt-v'                        = 'f338'
        'assistive-listening-systems'         = 'f2a2'
        'asterisk'                            = 'f069'
        'at'                                  = 'f1fa'
        'atlas'                               = 'f558'
        'atom'                                = 'f5d2'
        'audio-description'                   = 'f29e'
        'award'                               = 'f559'
        'baby'                                = 'f77c'
        'baby-carriage'                       = 'f77d'
        'backspace'                           = 'f55a'
        'backward'                            = 'f04a'
        'bacon'                               = 'f7e5'
        'balance-scale'                       = 'f24e'
        'balance-scale-left'                  = 'f515'
        'balance-scale-right'                 = 'f516'
        'ban'                                 = 'f05e'
        'band-aid'                            = 'f462'
        'barcode'                             = 'f02a'
        'bars'                                = 'f0c9'
        'baseball-ball'                       = 'f433'
        'basketball-ball'                     = 'f434'
        'bath'                                = 'f2cd'
        'battery-empty'                       = 'f244'
        'battery-full'                        = 'f240'
        'battery-half'                        = 'f242'
        'battery-quarter'                     = 'f243'
        'battery-three-quarters'              = 'f241'
        'bed'                                 = 'f236'
        'beer'                                = 'f0fc'
        'bell'                                = 'f0f3'
        'bell-slash'                          = 'f1f6'
        'bezier-curve'                        = 'f55b'
        'bible'                               = 'f647'
        'bicycle'                             = 'f206'
        'biking'                              = 'f84a'
        'binoculars'                          = 'f1e5'
        'biohazard'                           = 'f780'
        'birthday-cake'                       = 'f1fd'
        'blender'                             = 'f517'
        'blender-phone'                       = 'f6b6'
        'blind'                               = 'f29d'
        'blog'                                = 'f781'
        'bold'                                = 'f032'
        'bolt'                                = 'f0e7'
        'bomb'                                = 'f1e2'
        'bone'                                = 'f5d7'
        'bong'                                = 'f55c'
        'book'                                = 'f02d'
        'book-dead'                           = 'f6b7'
        'book-medical'                        = 'f7e6'
        'book-open'                           = 'f518'
        'book-reader'                         = 'f5da'
        'bookmark'                            = 'f02e'
        'border-all'                          = 'f84c'
        'border-none'                         = 'f850'
        'border-style'                        = 'f853'
        'bowling-ball'                        = 'f436'
        'box'                                 = 'f466'
        'box-open'                            = 'f49e'
        'boxes'                               = 'f468'
        'braille'                             = 'f2a1'
        'brain'                               = 'f5dc'
        'bread-slice'                         = 'f7ec'
        'briefcase'                           = 'f0b1'
        'briefcase-medical'                   = 'f469'
        'broadcast-tower'                     = 'f519'
        'broom'                               = 'f51a'
        'brush'                               = 'f55d'
        'bug'                                 = 'f188'
        'building'                            = 'f1ad'
        'bullhorn'                            = 'f0a1'
        'bullseye'                            = 'f140'
        'burn'                                = 'f46a'
        'bus'                                 = 'f207'
        'bus-alt'                             = 'f55e'
        'business-time'                       = 'f64a'
        'calculator'                          = 'f1ec'
        'calendar'                            = 'f133'
        'calendar-alt'                        = 'f073'
        'calendar-check'                      = 'f274'
        'calendar-day'                        = 'f783'
        'calendar-minus'                      = 'f272'
        'calendar-plus'                       = 'f271'
        'calendar-times'                      = 'f273'
        'calendar-week'                       = 'f784'
        'camera'                              = 'f030'
        'camera-retro'                        = 'f083'
        'campground'                          = 'f6bb'
        'candy-cane'                          = 'f786'
        'cannabis'                            = 'f55f'
        'capsules'                            = 'f46b'
        'car'                                 = 'f1b9'
        'car-alt'                             = 'f5de'
        'car-battery'                         = 'f5df'
        'car-crash'                           = 'f5e1'
        'car-side'                            = 'f5e4'
        'caret-down'                          = 'f0d7'
        'caret-left'                          = 'f0d9'
        'caret-right'                         = 'f0da'
        'caret-square-down'                   = 'f150'
        'caret-square-left'                   = 'f191'
        'caret-square-right'                  = 'f152'
        'caret-square-up'                     = 'f151'
        'caret-up'                            = 'f0d8'
        'carrot'                              = 'f787'
        'cart-arrow-down'                     = 'f218'
        'cart-plus'                           = 'f217'
        'cash-register'                       = 'f788'
        'cat'                                 = 'f6be'
        'certificate'                         = 'f0a3'
        'chair'                               = 'f6c0'
        'chalkboard'                          = 'f51b'
        'chalkboard-teacher'                  = 'f51c'
        'charging-station'                    = 'f5e7'
        'chart-area'                          = 'f1fe'
        'chart-bar'                           = 'f080'
        'chart-line'                          = 'f201'
        'chart-pie'                           = 'f200'
        'check'                               = 'f00c'
        'check-circle'                        = 'f058'
        'check-double'                        = 'f560'
        'check-square'                        = 'f14a'
        'cheese'                              = 'f7ef'
        'chess'                               = 'f439'
        'chess-bishop'                        = 'f43a'
        'chess-board'                         = 'f43c'
        'chess-king'                          = 'f43f'
        'chess-knight'                        = 'f441'
        'chess-pawn'                          = 'f443'
        'chess-queen'                         = 'f445'
        'chess-rook'                          = 'f447'
        'chevron-circle-down'                 = 'f13a'
        'chevron-circle-left'                 = 'f137'
        'chevron-circle-right'                = 'f138'
        'chevron-circle-up'                   = 'f139'
        'chevron-down'                        = 'f078'
        'chevron-left'                        = 'f053'
        'chevron-right'                       = 'f054'
        'chevron-up'                          = 'f077'
        'child'                               = 'f1ae'
        'church'                              = 'f51d'
        'circle'                              = 'f111'
        'circle-notch'                        = 'f1ce'
        'city'                                = 'f64f'
        'clinic-medical'                      = 'f7f2'
        'clipboard'                           = 'f328'
        'clipboard-check'                     = 'f46c'
        'clipboard-list'                      = 'f46d'
        'clock'                               = 'f017'
        'clone'                               = 'f24d'
        'closed-captioning'                   = 'f20a'
        'cloud'                               = 'f0c2'
        'cloud-download-alt'                  = 'f381'
        'cloud-meatball'                      = 'f73b'
        'cloud-moon'                          = 'f6c3'
        'cloud-moon-rain'                     = 'f73c'
        'cloud-rain'                          = 'f73d'
        'cloud-showers-heavy'                 = 'f740'
        'cloud-sun'                           = 'f6c4'
        'cloud-sun-rain'                      = 'f743'
        'cloud-upload-alt'                    = 'f382'
        'cocktail'                            = 'f561'
        'code'                                = 'f121'
        'code-branch'                         = 'f126'
        'coffee'                              = 'f0f4'
        'cog'                                 = 'f013'
        'cogs'                                = 'f085'
        'coins'                               = 'f51e'
        'columns'                             = 'f0db'
        'comment'                             = 'f075'
        'comment-alt'                         = 'f27a'
        'comment-dollar'                      = 'f651'
        'comment-dots'                        = 'f4ad'
        'comment-medical'                     = 'f7f5'
        'comment-slash'                       = 'f4b3'
        'comments'                            = 'f086'
        'comments-dollar'                     = 'f653'
        'compact-disc'                        = 'f51f'
        'compass'                             = 'f14e'
        'compress'                            = 'f066'
        'compress-arrows-alt'                 = 'f78c'
        'concierge-bell'                      = 'f562'
        'cookie'                              = 'f563'
        'cookie-bite'                         = 'f564'
        'copy'                                = 'f0c5'
        'copyright'                           = 'f1f9'
        'couch'                               = 'f4b8'
        'credit-card'                         = 'f09d'
        'crop'                                = 'f125'
        'crop-alt'                            = 'f565'
        'cross'                               = 'f654'
        'crosshairs'                          = 'f05b'
        'crow'                                = 'f520'
        'crown'                               = 'f521'
        'crutch'                              = 'f7f7'
        'cube'                                = 'f1b2'
        'cubes'                               = 'f1b3'
        'cut'                                 = 'f0c4'
        'database'                            = 'f1c0'
        'deaf'                                = 'f2a4'
        'democrat'                            = 'f747'
        'desktop'                             = 'f108'
        'dharmachakra'                        = 'f655'
        'diagnoses'                           = 'f470'
        'dice'                                = 'f522'
        'dice-d20'                            = 'f6cf'
        'dice-d6'                             = 'f6d1'
        'dice-five'                           = 'f523'
        'dice-four'                           = 'f524'
        'dice-one'                            = 'f525'
        'dice-six'                            = 'f526'
        'dice-three'                          = 'f527'
        'dice-two'                            = 'f528'
        'digital-tachograph'                  = 'f566'
        'directions'                          = 'f5eb'
        'divide'                              = 'f529'
        'dizzy'                               = 'f567'
        'dna'                                 = 'f471'
        'dog'                                 = 'f6d3'
        'dollar-sign'                         = 'f155'
        'dolly'                               = 'f472'
        'dolly-flatbed'                       = 'f474'
        'donate'                              = 'f4b9'
        'door-closed'                         = 'f52a'
        'door-open'                           = 'f52b'
        'dot-circle'                          = 'f192'
        'dove'                                = 'f4ba'
        'download'                            = 'f019'
        'drafting-compass'                    = 'f568'
        'dragon'                              = 'f6d5'
        'draw-polygon'                        = 'f5ee'
        'drum'                                = 'f569'
        'drum-steelpan'                       = 'f56a'
        'drumstick-bite'                      = 'f6d7'
        'dumbbell'                            = 'f44b'
        'dumpster'                            = 'f793'
        'dumpster-fire'                       = 'f794'
        'dungeon'                             = 'f6d9'
        'edit'                                = 'f044'
        'egg'                                 = 'f7fb'
        'eject'                               = 'f052'
        'ellipsis-h'                          = 'f141'
        'ellipsis-v'                          = 'f142'
        'envelope'                            = 'f0e0'
        'envelope-open'                       = 'f2b6'
        'envelope-open-text'                  = 'f658'
        'envelope-square'                     = 'f199'
        'equals'                              = 'f52c'
        'eraser'                              = 'f12d'
        'ethernet'                            = 'f796'
        'euro-sign'                           = 'f153'
        'exchange-alt'                        = 'f362'
        'exclamation'                         = 'f12a'
        'exclamation-circle'                  = 'f06a'
        'exclamation-triangle'                = 'f071'
        'expand'                              = 'f065'
        'expand-arrows-alt'                   = 'f31e'
        'external-link-alt'                   = 'f35d'
        'external-link-square-alt'            = 'f360'
        'eye'                                 = 'f06e'
        'eye-dropper'                         = 'f1fb'
        'eye-slash'                           = 'f070'
        'fan'                                 = 'f863'
        'fast-backward'                       = 'f049'
        'fast-forward'                        = 'f050'
        'fax'                                 = 'f1ac'
        'feather'                             = 'f52d'
        'feather-alt'                         = 'f56b'
        'female'                              = 'f182'
        'fighter-jet'                         = 'f0fb'
        'file'                                = 'f15b'
        'file-alt'                            = 'f15c'
        'file-archive'                        = 'f1c6'
        'file-audio'                          = 'f1c7'
        'file-code'                           = 'f1c9'
        'file-contract'                       = 'f56c'
        'file-csv'                            = 'f6dd'
        'file-download'                       = 'f56d'
        'file-excel'                          = 'f1c3'
        'file-export'                         = 'f56e'
        'file-image'                          = 'f1c5'
        'file-import'                         = 'f56f'
        'file-invoice'                        = 'f570'
        'file-invoice-dollar'                 = 'f571'
        'file-medical'                        = 'f477'
        'file-medical-alt'                    = 'f478'
        'file-pdf'                            = 'f1c1'
        'file-powerpoint'                     = 'f1c4'
        'file-prescription'                   = 'f572'
        'file-signature'                      = 'f573'
        'file-upload'                         = 'f574'
        'file-video'                          = 'f1c8'
        'file-word'                           = 'f1c2'
        'fill'                                = 'f575'
        'fill-drip'                           = 'f576'
        'film'                                = 'f008'
        'filter'                              = 'f0b0'
        'fingerprint'                         = 'f577'
        'fire'                                = 'f06d'
        'fire-alt'                            = 'f7e4'
        'fire-extinguisher'                   = 'f134'
        'first-aid'                           = 'f479'
        'fish'                                = 'f578'
        'fist-raised'                         = 'f6de'
        'flag'                                = 'f024'
        'flag-checkered'                      = 'f11e'
        'flag-usa'                            = 'f74d'
        'flask'                               = 'f0c3'
        'flushed'                             = 'f579'
        'folder'                              = 'f07b'
        'folder-minus'                        = 'f65d'
        'folder-open'                         = 'f07c'
        'folder-plus'                         = 'f65e'
        'font'                                = 'f031'
        'football-ball'                       = 'f44e'
        'forward'                             = 'f04e'
        'frog'                                = 'f52e'
        'frown'                               = 'f119'
        'frown-open'                          = 'f57a'
        'funnel-dollar'                       = 'f662'
        'futbol'                              = 'f1e3'
        'gamepad'                             = 'f11b'
        'gas-pump'                            = 'f52f'
        'gavel'                               = 'f0e3'
        'gem'                                 = 'f3a5'
        'genderless'                          = 'f22d'
        'ghost'                               = 'f6e2'
        'gift'                                = 'f06b'
        'gifts'                               = 'f79c'
        'glass-cheers'                        = 'f79f'
        'glass-martini'                       = 'f000'
        'glass-martini-alt'                   = 'f57b'
        'glass-whiskey'                       = 'f7a0'
        'glasses'                             = 'f530'
        'globe'                               = 'f0ac'
        'globe-africa'                        = 'f57c'
        'globe-americas'                      = 'f57d'
        'globe-asia'                          = 'f57e'
        'globe-europe'                        = 'f7a2'
        'golf-ball'                           = 'f450'
        'gopuram'                             = 'f664'
        'graduation-cap'                      = 'f19d'
        'greater-than'                        = 'f531'
        'greater-than-equal'                  = 'f532'
        'grimace'                             = 'f57f'
        'grin'                                = 'f580'
        'grin-alt'                            = 'f581'
        'grin-beam'                           = 'f582'
        'grin-beam-sweat'                     = 'f583'
        'grin-hearts'                         = 'f584'
        'grin-squint'                         = 'f585'
        'grin-squint-tears'                   = 'f586'
        'grin-stars'                          = 'f587'
        'grin-tears'                          = 'f588'
        'grin-tongue'                         = 'f589'
        'grin-tongue-squint'                  = 'f58a'
        'grin-tongue-wink'                    = 'f58b'
        'grin-wink'                           = 'f58c'
        'grip-horizontal'                     = 'f58d'
        'grip-lines'                          = 'f7a4'
        'grip-lines-vertical'                 = 'f7a5'
        'grip-vertical'                       = 'f58e'
        'guitar'                              = 'f7a6'
        'h-square'                            = 'f0fd'
        'hamburger'                           = 'f805'
        'hammer'                              = 'f6e3'
        'hamsa'                               = 'f665'
        'hand-holding'                        = 'f4bd'
        'hand-holding-heart'                  = 'f4be'
        'hand-holding-usd'                    = 'f4c0'
        'hand-lizard'                         = 'f258'
        'hand-middle-finger'                  = 'f806'
        'hand-paper'                          = 'f256'
        'hand-peace'                          = 'f25b'
        'hand-point-down'                     = 'f0a7'
        'hand-point-left'                     = 'f0a5'
        'hand-point-right'                    = 'f0a4'
        'hand-point-up'                       = 'f0a6'
        'hand-pointer'                        = 'f25a'
        'hand-rock'                           = 'f255'
        'hand-scissors'                       = 'f257'
        'hand-spock'                          = 'f259'
        'hands'                               = 'f4c2'
        'hands-helping'                       = 'f4c4'
        'handshake'                           = 'f2b5'
        'hanukiah'                            = 'f6e6'
        'hard-hat'                            = 'f807'
        'hashtag'                             = 'f292'
        'hat-cowboy'                          = 'f8c0'
        'hat-cowboy-side'                     = 'f8c1'
        'hat-wizard'                          = 'f6e8'
        'haykal'                              = 'f666'
        'hdd'                                 = 'f0a0'
        'heading'                             = 'f1dc'
        'headphones'                          = 'f025'
        'headphones-alt'                      = 'f58f'
        'headset'                             = 'f590'
        'heart'                               = 'f004'
        'heart-broken'                        = 'f7a9'
        'heartbeat'                           = 'f21e'
        'helicopter'                          = 'f533'
        'highlighter'                         = 'f591'
        'hiking'                              = 'f6ec'
        'hippo'                               = 'f6ed'
        'history'                             = 'f1da'
        'hockey-puck'                         = 'f453'
        'holly-berry'                         = 'f7aa'
        'home'                                = 'f015'
        'horse'                               = 'f6f0'
        'horse-head'                          = 'f7ab'
        'hospital'                            = 'f0f8'
        'hospital-alt'                        = 'f47d'
        'hospital-symbol'                     = 'f47e'
        'hot-tub'                             = 'f593'
        'hotdog'                              = 'f80f'
        'hotel'                               = 'f594'
        'hourglass'                           = 'f254'
        'hourglass-end'                       = 'f253'
        'hourglass-half'                      = 'f252'
        'hourglass-start'                     = 'f251'
        'house-damage'                        = 'f6f1'
        'hryvnia'                             = 'f6f2'
        'i-cursor'                            = 'f246'
        'ice-cream'                           = 'f810'
        'icicles'                             = 'f7ad'
        'icons'                               = 'f86d'
        'id-badge'                            = 'f2c1'
        'id-card'                             = 'f2c2'
        'id-card-alt'                         = 'f47f'
        'igloo'                               = 'f7ae'
        'image'                               = 'f03e'
        'images'                              = 'f302'
        'inbox'                               = 'f01c'
        'indent'                              = 'f03c'
        'industry'                            = 'f275'
        'infinity'                            = 'f534'
        'info'                                = 'f129'
        'info-circle'                         = 'f05a'
        'italic'                              = 'f033'
        'jedi'                                = 'f669'
        'joint'                               = 'f595'
        'journal-whills'                      = 'f66a'
        'kaaba'                               = 'f66b'
        'key'                                 = 'f084'
        'keyboard'                            = 'f11c'
        'khanda'                              = 'f66d'
        'kiss'                                = 'f596'
        'kiss-beam'                           = 'f597'
        'kiss-wink-heart'                     = 'f598'
        'kiwi-bird'                           = 'f535'
        'landmark'                            = 'f66f'
        'language'                            = 'f1ab'
        'laptop'                              = 'f109'
        'laptop-code'                         = 'f5fc'
        'laptop-medical'                      = 'f812'
        'laugh'                               = 'f599'
        'laugh-beam'                          = 'f59a'
        'laugh-squint'                        = 'f59b'
        'laugh-wink'                          = 'f59c'
        'layer-group'                         = 'f5fd'
        'leaf'                                = 'f06c'
        'lemon'                               = 'f094'
        'less-than'                           = 'f536'
        'less-than-equal'                     = 'f537'
        'level-down-alt'                      = 'f3be'
        'level-up-alt'                        = 'f3bf'
        'life-ring'                           = 'f1cd'
        'lightbulb'                           = 'f0eb'
        'link'                                = 'f0c1'
        'lira-sign'                           = 'f195'
        'list'                                = 'f03a'
        'list-alt'                            = 'f022'
        'list-ol'                             = 'f0cb'
        'list-ul'                             = 'f0ca'
        'location-arrow'                      = 'f124'
        'lock'                                = 'f023'
        'lock-open'                           = 'f3c1'
        'long-arrow-alt-down'                 = 'f309'
        'long-arrow-alt-left'                 = 'f30a'
        'long-arrow-alt-right'                = 'f30b'
        'long-arrow-alt-up'                   = 'f30c'
        'low-vision'                          = 'f2a8'
        'luggage-cart'                        = 'f59d'
        'magic'                               = 'f0d0'
        'magnet'                              = 'f076'
        'mail-bulk'                           = 'f674'
        'male'                                = 'f183'
        'map'                                 = 'f279'
        'map-marked'                          = 'f59f'
        'map-marked-alt'                      = 'f5a0'
        'map-marker'                          = 'f041'
        'map-marker-alt'                      = 'f3c5'
        'map-pin'                             = 'f276'
        'map-signs'                           = 'f277'
        'marker'                              = 'f5a1'
        'mars'                                = 'f222'
        'mars-double'                         = 'f227'
        'mars-stroke'                         = 'f229'
        'mars-stroke-h'                       = 'f22b'
        'mars-stroke-v'                       = 'f22a'
        'mask'                                = 'f6fa'
        'medal'                               = 'f5a2'
        'medkit'                              = 'f0fa'
        'meh'                                 = 'f11a'
        'meh-blank'                           = 'f5a4'
        'meh-rolling-eyes'                    = 'f5a5'
        'memory'                              = 'f538'
        'menorah'                             = 'f676'
        'mercury'                             = 'f223'
        'meteor'                              = 'f753'
        'microchip'                           = 'f2db'
        'microphone'                          = 'f130'
        'microphone-alt'                      = 'f3c9'
        'microphone-alt-slash'                = 'f539'
        'microphone-slash'                    = 'f131'
        'microscope'                          = 'f610'
        'minus'                               = 'f068'
        'minus-circle'                        = 'f056'
        'minus-square'                        = 'f146'
        'mitten'                              = 'f7b5'
        'mobile'                              = 'f10b'
        'mobile-alt'                          = 'f3cd'
        'money-bill'                          = 'f0d6'
        'money-bill-alt'                      = 'f3d1'
        'money-bill-wave'                     = 'f53a'
        'money-bill-wave-alt'                 = 'f53b'
        'money-check'                         = 'f53c'
        'money-check-alt'                     = 'f53d'
        'monument'                            = 'f5a6'
        'moon'                                = 'f186'
        'mortar-pestle'                       = 'f5a7'
        'mosque'                              = 'f678'
        'motorcycle'                          = 'f21c'
        'mountain'                            = 'f6fc'
        'mouse'                               = 'f8cc'
        'mouse-pointer'                       = 'f245'
        'mug-hot'                             = 'f7b6'
        'music'                               = 'f001'
        'network-wired'                       = 'f6ff'
        'neuter'                              = 'f22c'
        'newspaper'                           = 'f1ea'
        'not-equal'                           = 'f53e'
        'notes-medical'                       = 'f481'
        'object-group'                        = 'f247'
        'object-ungroup'                      = 'f248'
        'oil-can'                             = 'f613'
        'om'                                  = 'f679'
        'otter'                               = 'f700'
        'outdent'                             = 'f03b'
        'pager'                               = 'f815'
        'paint-brush'                         = 'f1fc'
        'paint-roller'                        = 'f5aa'
        'palette'                             = 'f53f'
        'pallet'                              = 'f482'
        'paper-plane'                         = 'f1d8'
        'paperclip'                           = 'f0c6'
        'parachute-box'                       = 'f4cd'
        'paragraph'                           = 'f1dd'
        'parking'                             = 'f540'
        'passport'                            = 'f5ab'
        'pastafarianism'                      = 'f67b'
        'paste'                               = 'f0ea'
        'pause'                               = 'f04c'
        'pause-circle'                        = 'f28b'
        'paw'                                 = 'f1b0'
        'peace'                               = 'f67c'
        'pen'                                 = 'f304'
        'pen-alt'                             = 'f305'
        'pen-fancy'                           = 'f5ac'
        'pen-nib'                             = 'f5ad'
        'pen-square'                          = 'f14b'
        'pencil-alt'                          = 'f303'
        'pencil-ruler'                        = 'f5ae'
        'people-carry'                        = 'f4ce'
        'pepper-hot'                          = 'f816'
        'percent'                             = 'f295'
        'percentage'                          = 'f541'
        'person-booth'                        = 'f756'
        'phone'                               = 'f095'
        'phone-alt'                           = 'f879'
        'phone-slash'                         = 'f3dd'
        'phone-square'                        = 'f098'
        'phone-square-alt'                    = 'f87b'
        'phone-volume'                        = 'f2a0'
        'photo-video'                         = 'f87c'
        'piggy-bank'                          = 'f4d3'
        'pills'                               = 'f484'
        'pizza-slice'                         = 'f818'
        'place-of-worship'                    = 'f67f'
        'plane'                               = 'f072'
        'plane-arrival'                       = 'f5af'
        'plane-departure'                     = 'f5b0'
        'play'                                = 'f04b'
        'play-circle'                         = 'f144'
        'plug'                                = 'f1e6'
        'plus'                                = 'f067'
        'plus-circle'                         = 'f055'
        'plus-square'                         = 'f0fe'
        'podcast'                             = 'f2ce'
        'poll'                                = 'f681'
        'poll-h'                              = 'f682'
        'poo'                                 = 'f2fe'
        'poo-storm'                           = 'f75a'
        'poop'                                = 'f619'
        'portrait'                            = 'f3e0'
        'pound-sign'                          = 'f154'
        'power-off'                           = 'f011'
        'pray'                                = 'f683'
        'praying-hands'                       = 'f684'
        'prescription'                        = 'f5b1'
        'prescription-bottle'                 = 'f485'
        'prescription-bottle-alt'             = 'f486'
        'print'                               = 'f02f'
        'procedures'                          = 'f487'
        'project-diagram'                     = 'f542'
        'puzzle-piece'                        = 'f12e'
        'qrcode'                              = 'f029'
        'question'                            = 'f128'
        'question-circle'                     = 'f059'
        'quidditch'                           = 'f458'
        'quote-left'                          = 'f10d'
        'quote-right'                         = 'f10e'
        'quran'                               = 'f687'
        'radiation'                           = 'f7b9'
        'radiation-alt'                       = 'f7ba'
        'rainbow'                             = 'f75b'
        'random'                              = 'f074'
        'receipt'                             = 'f543'
        'record-vinyl'                        = 'f8d9'
        'recycle'                             = 'f1b8'
        'redo'                                = 'f01e'
        'redo-alt'                            = 'f2f9'
        'registered'                          = 'f25d'
        'remove-format'                       = 'f87d'
        'reply'                               = 'f3e5'
        'reply-all'                           = 'f122'
        'republican'                          = 'f75e'
        'restroom'                            = 'f7bd'
        'retweet'                             = 'f079'
        'ribbon'                              = 'f4d6'
        'ring'                                = 'f70b'
        'road'                                = 'f018'
        'robot'                               = 'f544'
        'rocket'                              = 'f135'
        'route'                               = 'f4d7'
        'rss'                                 = 'f09e'
        'rss-square'                          = 'f143'
        'ruble-sign'                          = 'f158'
        'ruler'                               = 'f545'
        'ruler-combined'                      = 'f546'
        'ruler-horizontal'                    = 'f547'
        'ruler-vertical'                      = 'f548'
        'running'                             = 'f70c'
        'rupee-sign'                          = 'f156'
        'sad-cry'                             = 'f5b3'
        'sad-tear'                            = 'f5b4'
        'satellite'                           = 'f7bf'
        'satellite-dish'                      = 'f7c0'
        'save'                                = 'f0c7'
        'school'                              = 'f549'
        'screwdriver'                         = 'f54a'
        'scroll'                              = 'f70e'
        'sd-card'                             = 'f7c2'
        'search'                              = 'f002'
        'search-dollar'                       = 'f688'
        'search-location'                     = 'f689'
        'search-minus'                        = 'f010'
        'search-plus'                         = 'f00e'
        'seedling'                            = 'f4d8'
        'server'                              = 'f233'
        'shapes'                              = 'f61f'
        'share'                               = 'f064'
        'share-alt'                           = 'f1e0'
        'share-alt-square'                    = 'f1e1'
        'share-square'                        = 'f14d'
        'shekel-sign'                         = 'f20b'
        'shield-alt'                          = 'f3ed'
        'ship'                                = 'f21a'
        'shipping-fast'                       = 'f48b'
        'shoe-prints'                         = 'f54b'
        'shopping-bag'                        = 'f290'
        'shopping-basket'                     = 'f291'
        'shopping-cart'                       = 'f07a'
        'shower'                              = 'f2cc'
        'shuttle-van'                         = 'f5b6'
        'sign'                                = 'f4d9'
        'sign-in-alt'                         = 'f2f6'
        'sign-language'                       = 'f2a7'
        'sign-out-alt'                        = 'f2f5'
        'signal'                              = 'f012'
        'signature'                           = 'f5b7'
        'sim-card'                            = 'f7c4'
        'sitemap'                             = 'f0e8'
        'skating'                             = 'f7c5'
        'skiing'                              = 'f7c9'
        'skiing-nordic'                       = 'f7ca'
        'skull'                               = 'f54c'
        'skull-crossbones'                    = 'f714'
        'slash'                               = 'f715'
        'sleigh'                              = 'f7cc'
        'sliders-h'                           = 'f1de'
        'smile'                               = 'f118'
        'smile-beam'                          = 'f5b8'
        'smile-wink'                          = 'f4da'
        'smog'                                = 'f75f'
        'smoking'                             = 'f48d'
        'smoking-ban'                         = 'f54d'
        'sms'                                 = 'f7cd'
        'snowboarding'                        = 'f7ce'
        'snowflake'                           = 'f2dc'
        'snowman'                             = 'f7d0'
        'snowplow'                            = 'f7d2'
        'socks'                               = 'f696'
        'solar-panel'                         = 'f5ba'
        'sort'                                = 'f0dc'
        'sort-alpha-down'                     = 'f15d'
        'sort-alpha-down-alt'                 = 'f881'
        'sort-alpha-up'                       = 'f15e'
        'sort-alpha-up-alt'                   = 'f882'
        'sort-amount-down'                    = 'f160'
        'sort-amount-down-alt'                = 'f884'
        'sort-amount-up'                      = 'f161'
        'sort-amount-up-alt'                  = 'f885'
        'sort-down'                           = 'f0dd'
        'sort-numeric-down'                   = 'f162'
        'sort-numeric-down-alt'               = 'f886'
        'sort-numeric-up'                     = 'f163'
        'sort-numeric-up-alt'                 = 'f887'
        'sort-up'                             = 'f0de'
        'spa'                                 = 'f5bb'
        'space-shuttle'                       = 'f197'
        'spell-check'                         = 'f891'
        'spider'                              = 'f717'
        'spinner'                             = 'f110'
        'splotch'                             = 'f5bc'
        'spray-can'                           = 'f5bd'
        'square'                              = 'f0c8'
        'square-full'                         = 'f45c'
        'square-root-alt'                     = 'f698'
        'stamp'                               = 'f5bf'
        'star'                                = 'f005'
        'star-and-crescent'                   = 'f699'
        'star-half'                           = 'f089'
        'star-half-alt'                       = 'f5c0'
        'star-of-david'                       = 'f69a'
        'star-of-life'                        = 'f621'
        'step-backward'                       = 'f048'
        'step-forward'                        = 'f051'
        'stethoscope'                         = 'f0f1'
        'sticky-note'                         = 'f249'
        'stop'                                = 'f04d'
        'stop-circle'                         = 'f28d'
        'stopwatch'                           = 'f2f2'
        'store'                               = 'f54e'
        'store-alt'                           = 'f54f'
        'stream'                              = 'f550'
        'street-view'                         = 'f21d'
        'strikethrough'                       = 'f0cc'
        'stroopwafel'                         = 'f551'
        'subscript'                           = 'f12c'
        'subway'                              = 'f239'
        'suitcase'                            = 'f0f2'
        'suitcase-rolling'                    = 'f5c1'
        'sun'                                 = 'f185'
        'superscript'                         = 'f12b'
        'surprise'                            = 'f5c2'
        'swatchbook'                          = 'f5c3'
        'swimmer'                             = 'f5c4'
        'swimming-pool'                       = 'f5c5'
        'synagogue'                           = 'f69b'
        'sync'                                = 'f021'
        'sync-alt'                            = 'f2f1'
        'syringe'                             = 'f48e'
        'table'                               = 'f0ce'
        'table-tennis'                        = 'f45d'
        'tablet'                              = 'f10a'
        'tablet-alt'                          = 'f3fa'
        'tablets'                             = 'f490'
        'tachometer-alt'                      = 'f3fd'
        'tag'                                 = 'f02b'
        'tags'                                = 'f02c'
        'tape'                                = 'f4db'
        'tasks'                               = 'f0ae'
        'taxi'                                = 'f1ba'
        'teeth'                               = 'f62e'
        'teeth-open'                          = 'f62f'
        'temperature-high'                    = 'f769'
        'temperature-low'                     = 'f76b'
        'tenge'                               = 'f7d7'
        'terminal'                            = 'f120'
        'text-height'                         = 'f034'
        'text-width'                          = 'f035'
        'th'                                  = 'f00a'
        'th-large'                            = 'f009'
        'th-list'                             = 'f00b'
        'theater-masks'                       = 'f630'
        'thermometer'                         = 'f491'
        'thermometer-empty'                   = 'f2cb'
        'thermometer-full'                    = 'f2c7'
        'thermometer-half'                    = 'f2c9'
        'thermometer-quarter'                 = 'f2ca'
        'thermometer-three-quarters'          = 'f2c8'
        'thumbs-down'                         = 'f165'
        'thumbs-up'                           = 'f164'
        'thumbtack'                           = 'f08d'
        'ticket-alt'                          = 'f3ff'
        'times'                               = 'f00d'
        'times-circle'                        = 'f057'
        'tint'                                = 'f043'
        'tint-slash'                          = 'f5c7'
        'tired'                               = 'f5c8'
        'toggle-off'                          = 'f204'
        'toggle-on'                           = 'f205'
        'toilet'                              = 'f7d8'
        'toilet-paper'                        = 'f71e'
        'toolbox'                             = 'f552'
        'tools'                               = 'f7d9'
        'tooth'                               = 'f5c9'
        'torah'                               = 'f6a0'
        'torii-gate'                          = 'f6a1'
        'tractor'                             = 'f722'
        'trademark'                           = 'f25c'
        'traffic-light'                       = 'f637'
        'train'                               = 'f238'
        'tram'                                = 'f7da'
        'transgender'                         = 'f224'
        'transgender-alt'                     = 'f225'
        'trash'                               = 'f1f8'
        'trash-alt'                           = 'f2ed'
        'trash-restore'                       = 'f829'
        'trash-restore-alt'                   = 'f82a'
        'tree'                                = 'f1bb'
        'trophy'                              = 'f091'
        'truck'                               = 'f0d1'
        'truck-loading'                       = 'f4de'
        'truck-monster'                       = 'f63b'
        'truck-moving'                        = 'f4df'
        'truck-pickup'                        = 'f63c'
        'tshirt'                              = 'f553'
        'tty'                                 = 'f1e4'
        'tv'                                  = 'f26c'
        'umbrella'                            = 'f0e9'
        'umbrella-beach'                      = 'f5ca'
        'underline'                           = 'f0cd'
        'undo'                                = 'f0e2'
        'undo-alt'                            = 'f2ea'
        'universal-access'                    = 'f29a'
        'university'                          = 'f19c'
        'unlink'                              = 'f127'
        'unlock'                              = 'f09c'
        'unlock-alt'                          = 'f13e'
        'upload'                              = 'f093'
        'user'                                = 'f007'
        'user-alt'                            = 'f406'
        'user-alt-slash'                      = 'f4fa'
        'user-astronaut'                      = 'f4fb'
        'user-check'                          = 'f4fc'
        'user-circle'                         = 'f2bd'
        'user-clock'                          = 'f4fd'
        'user-cog'                            = 'f4fe'
        'user-edit'                           = 'f4ff'
        'user-friends'                        = 'f500'
        'user-graduate'                       = 'f501'
        'user-injured'                        = 'f728'
        'user-lock'                           = 'f502'
        'user-md'                             = 'f0f0'
        'user-minus'                          = 'f503'
        'user-ninja'                          = 'f504'
        'user-nurse'                          = 'f82f'
        'user-plus'                           = 'f234'
        'user-secret'                         = 'f21b'
        'user-shield'                         = 'f505'
        'user-slash'                          = 'f506'
        'user-tag'                            = 'f507'
        'user-tie'                            = 'f508'
        'user-times'                          = 'f235'
        'users'                               = 'f0c0'
        'users-cog'                           = 'f509'
        'utensil-spoon'                       = 'f2e5'
        'utensils'                            = 'f2e7'
        'vector-square'                       = 'f5cb'
        'venus'                               = 'f221'
        'venus-double'                        = 'f226'
        'venus-mars'                          = 'f228'
        'vial'                                = 'f492'
        'vials'                               = 'f493'
        'video'                               = 'f03d'
        'video-slash'                         = 'f4e2'
        'vihara'                              = 'f6a7'
        'voicemail'                           = 'f897'
        'volleyball-ball'                     = 'f45f'
        'volume-down'                         = 'f027'
        'volume-mute'                         = 'f6a9'
        'volume-off'                          = 'f026'
        'volume-up'                           = 'f028'
        'vote-yea'                            = 'f772'
        'vr-cardboard'                        = 'f729'
        'walking'                             = 'f554'
        'wallet'                              = 'f555'
        'warehouse'                           = 'f494'
        'water'                               = 'f773'
        'wave-square'                         = 'f83e'
        'weight'                              = 'f496'
        'weight-hanging'                      = 'f5cd'
        'wheelchair'                          = 'f193'
        'wifi'                                = 'f1eb'
        'wind'                                = 'f72e'
        'window-close'                        = 'f410'
        'window-maximize'                     = 'f2d0'
        'window-minimize'                     = 'f2d1'
        'window-restore'                      = 'f2d2'
        'wine-bottle'                         = 'f72f'
        'wine-glass'                          = 'f4e3'
        'wine-glass-alt'                      = 'f5ce'
        'won-sign'                            = 'f159'
        'wrench'                              = 'f0ad'
        'x-ray'                               = 'f497'
        'yen-sign'                            = 'f157'
        'yin-yang'                            = 'f6ad'
    }
}

function Set-Tag {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $HtmlObject,
        [switch] $NewLine # This is needed if code requires new lines such as JavaScript
    )
    $HTML = [System.Text.StringBuilder]::new()
    [void] $HTML.Append("<$($HtmlObject.Tag)")
    foreach ($Property in $HtmlObject.Attributes.Keys) {
        $PropertyValue = $HtmlObject.Attributes[$Property]
        # This checks if property has any subproperties  such as style having multiple options
        if ($PropertyValue -is [System.Collections.IDictionary]) {
            $OutputSubProperties = foreach ($SubAttributes in $PropertyValue.Keys) {
                $SubPropertyValue = $PropertyValue[$SubAttributes]
                # skip adding properties that are empty
                if ($null -ne $SubPropertyValue -and $SubPropertyValue -ne '') {
                    "$($SubAttributes):$($SubPropertyValue)"
                }
            }
            $MyValue = $OutputSubProperties -join ';'
            if ($MyValue.Trim() -ne '') {
                [void] $HTML.Append(" $Property=`"$MyValue`"")
            }
        } else {
            # skip adding properties that are empty
            if ($null -ne $PropertyValue -and $PropertyValue -ne '') {
                [void] $HTML.Append(" $Property=`"$PropertyValue`"")
            }
        }
    }
    if (($null -ne $HtmlObject.Value) -and ($HtmlObject.Value -ne '')) {
        [void] $HTML.Append(">")

        if ($HtmlObject.Value.Count -eq 1) {
            if ($HtmlObject.Value -is [System.Collections.IDictionary]) {
                [string] $NewObject = Set-Tag -HtmlObject ($HtmlObject.Value)
                [void] $HTML.Append($NewObject)
            } else {
                [void] $HTML.Append([string] $HtmlObject.Value)
            }
        } else {
            foreach ($Entry in $HtmlObject.Value) {
                if ($Entry -is [System.Collections.IDictionary]) {
                    [string] $NewObject = Set-Tag -HtmlObject ($Entry)
                    [void] $HTML.Append($NewObject)
                } else {
                    # This is needed if code requires new lines such as JavaScript
                    if ($NewLine) {
                        [void] $HTML.AppendLine([string] $Entry)
                    } else {
                        [void] $HTML.Append([string] $Entry)
                    }
                }
            }
        }
        [void] $HTML.Append("</$($HtmlObject.Tag)>")
    } else {
        if ($HtmlObject.SelfClosing) {
            [void] $HTML.Append("/>")
        } else {
            [void] $HTML.Append("></$($HtmlObject.Tag)>")
        }
    }
    $HTML.ToString()
}
function Email {
    [CmdLetBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)][ScriptBlock] $Email,
        [string[]] $To,
        [string[]] $CC,
        [string[]] $BCC,
        [string] $ReplyTo,
        [string] $From,
        [string] $Subject,
        [alias('SelfAttach')][switch] $AttachSelf,
        [string] $AttachSelfName,
        [string] $Server,
        [string] $Username,
        [int] $Port = 587,
        [string] $Password,
        [switch] $PasswordFromFile,
        [switch] $PasswordAsSecure,
        [switch] $SSL,
        [ValidateSet('Low', 'Normal', 'High')] [string] $Priority = 'Normal',
        [ValidateSet('None', 'OnSuccess', 'OnFailure', 'Delay', 'Never')] $DeliveryNotifications = 'None',
        [string] $Encoding = 'Unicode',
        [string] $FilePath,
        [bool] $Supress = $true,
        [switch] $WhatIf
    )
    $StartTime = [System.Diagnostics.Stopwatch]::StartNew()
    $ServerParameters = [ordered] @{
        From                  = $From
        To                    = $To
        CC                    = $CC
        BCC                   = $BCC
        ReplyTo               = $ReplyTo
        Server                = $Server
        Login                 = $Username
        Password              = $Password
        PasswordAsSecure      = $PasswordAsSecure
        PasswordFromFile      = $PasswordFromFile
        Port                  = $Port

        EnableSSL             = $SSL
        Encoding              = $Encoding
        Subject               = $Subject
        Priority              = $Priority
        DeliveryNotifications = $DeliveryNotifications
    }
    $Attachments = [System.Collections.Generic.List[string]]::new()
    $Body = New-HTML -UseCssLinks -UseJavaScriptLinks {
        [Array] $EmailParameters = Invoke-Command -ScriptBlock $Email

        foreach ($Parameter in $EmailParameters) {
            switch ( $Parameter.Type ) {
                HeaderTo {
                    $ServerParameters.To = $Parameter.Addresses
                }
                HeaderCC {
                    $ServerParameters.CC = $Parameter.Addresses
                }
                HeaderBCC {
                    $ServerParameters.BCC = $Parameter.Addresses
                }
                HeaderFrom {
                    $ServerParameters.From = $Parameter.Address
                }
                HeaderReplyTo {
                    $ServerParameters.ReplyTo = $Parameter.Address
                }
                HeaderSubject {
                    $ServerParameters.Subject = $Parameter.Subject
                }
                HeaderServer {
                    $ServerParameters.Server = $Parameter.Server
                    $ServerParameters.Port = $Parameter.Port
                    $ServerParameters.Login = $Parameter.UserName
                    $ServerParameters.Password = $Parameter.Password
                    $ServerParameters.PasswordFromFile = $Parameter.PasswordFromFile
                    $ServerParameters.PasswordAsSecure = $Parameter.PasswordAsSecure
                    $ServerParameters.EnableSSL = $Parameter.SSL
                }
                HeaderAttachment {
                    foreach ($Attachment in  $Parameter.FilePath) {
                        $Attachments.Add($Attachment)
                    }
                }
                HeaderOptions {
                    $ServerParameters.DeliveryNotifications = $Parameter.DeliveryNotifications
                    $ServerParameters.Encoding = $Parameter.Encoding
                    $ServerParameters.Priority = $Parameter.Priority
                }
                Default {
                    $Parameter
                }
            }
        }
    }
    if ($FilePath) {
        Save-HTML -FilePath $FilePath -HTML $Body
    }
    if ($AttachSelf) {
        if ($AttachSelfName) {
            $TempFilePath = "$(Get-TemporaryDirectory)\$($AttachSelfName).html"
        } else {
            $TempFilePath = ''
        }
        $Saved = Save-HTML -FilePath $TempFilePath -HTML $Body -Supress $false
        if ($Saved) {
            $Attachments.Add($Saved)
        }
    }

    #$MailSentTo = "To: $($ServerParameters.To -join ', '); CC: $($ServerParameters.CC -join ', '); BCC: $($ServerParameters.BCC -join ', ')".Trim()
    $EmailOutput = Send-Email -EmailParameters $ServerParameters -Body ($Body -join '') -Attachment $Attachments -WhatIf:$WhatIf
    if (-not $Supress) {
        $EmailOutput
    }

    $EndTime = Stop-TimeLog -Time $StartTime -Option OneLiner
    Write-Verbose "Email - Time to send: $EndTime"
}
function EmailAttachment {
    [CmdletBinding()]
    param(
        [string[]] $FilePath
    )
    [PSCustomObject] @{
        Type     = 'HeaderAttachment'
        FilePath = $FilePath
    }
}
function EmailBCC {
    [CmdletBinding()]
    param(
        [string[]] $Addresses
    )

    [PsCustomObject] @{
        Type      = 'HeaderBCC'
        Addresses = $Addresses
    }
}
function EmailBody {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)][ScriptBlock] $EmailBody,
        [string] $Color,
        [string] $BackGroundColor,
        [alias('Size')][int] $FontSize,
        [ValidateSet('normal', 'bold', 'bolder', 'lighter', '100', '200', '300', '400', '500', '600', '700', '800', '900')][string] $FontWeight,
        [ValidateSet('normal', 'italic', 'oblique')][string] $FontStyle,
        [ValidateSet('normal', 'small-caps')][string] $FontVariant,
        [string] $FontFamily ,
        [ValidateSet('left', 'center', 'right', 'justify')][string] $Alignment,
        [ValidateSet('none', 'line-through', 'overline', 'underline')][string] $TextDecoration,
        [ValidateSet('uppercase', 'lowercase', 'capitalize')][string] $TextTransform,
        [ValidateSet('rtl')][string] $Direction
    )

    $newHTMLSplat = @{ }
    if ($Alignment) {
        $newHTMLSplat.Alignment = $Alignment
    }
    if ($FontSize) {
        $newHTMLSplat.FontSize = $FontSize
    }
    if ($TextTransform) {
        $newHTMLSplat.TextTransform = $TextTransform
    }
    if ($Color) {
        $newHTMLSplat.Color = $Color
    }
    if ($FontFamily) {
        $newHTMLSplat.FontFamily = $FontFamily
    }
    if ($Direction) {
        $newHTMLSplat.Direction = $Direction
    }
    if ($FontStyle) {
        $newHTMLSplat.FontStyle = $FontStyle
    }
    if ($TextDecoration) {
        $newHTMLSplat.TextDecoration = $TextDecoration
    }
    if ($BackGroundColor) {
        $newHTMLSplat.BackGroundColor = $BackGroundColor
    }
    if ($FontVariant) {
        $newHTMLSplat.FontVariant = $FontVariant
    }
    if ($FontWeight) {
        $newHTMLSplat.FontWeight = $FontWeight
    }
    <#
    [bool] $SpanRequired = $false
    foreach ($Entry in $newHTMLSplat.GetEnumerator()) {
        if (($Entry.Value | Measure-Object).Count -gt 0) {
            $SpanRequired = $true
            break
        }
    }
    #>
    if ($newHTMLSplat.Count -gt 0) {
        $SpanRequired = $true
    } else {
        $SpanRequired = $false
    }
    if ($SpanRequired) {
        New-HTMLSpanStyle @newHTMLSplat {
            Invoke-Command -ScriptBlock $EmailBody
        }
    } else {
        Invoke-Command -ScriptBlock $EmailBody
    }
}

Register-ArgumentCompleter -CommandName EmailBody -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName EmailBody -ParameterName BackgroundColor -ScriptBlock { $Script:RGBColors.Keys }
function EmailCC {
    [CmdletBinding()]
    param(
        [string[]] $Addresses
    )

    [PsCustomObject] @{
        Type      = 'HeaderCC'
        Addresses = $Addresses
    }
}
function EmailFrom {
    [CmdletBinding()]
    param(
        [string] $Address
    )

    [PsCustomObject] @{
        Type    = 'HeaderFrom'
        Address = $Address
    }
}
function EmailHeader {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)][ScriptBlock] $EmailHeader
    )
    $EmailHeaders = Invoke-Command -ScriptBlock $EmailHeader
    $EmailHeaders
}
function EmailHTML {
    [CmdletBinding()]
    param(
        [ScriptBlock] $HTML
    )
    Invoke-Command -ScriptBlock $HTML
}
function EmailListItem {
    [CmdletBinding()]
    param(
        [string[]] $Text,
        [string[]] $Color = @(),
        [string[]] $BackGroundColor = @(),
        [int[]] $FontSize = @(),
        [ValidateSet('normal', 'bold', 'bolder', 'lighter', '100', '200', '300', '400', '500', '600', '700', '800', '900')][string[]] $FontWeight = @(),
        [ValidateSet('normal', 'italic', 'oblique')][string[]] $FontStyle = @(),
        [ValidateSet('normal', 'small-caps')][string[]] $FontVariant = @(),
        [string[]] $FontFamily = @(),
        [ValidateSet('left', 'center', 'right', 'justify')][string[]] $Alignment = @(),
        [ValidateSet('none', 'line-through', 'overline', 'underline')][string[]] $TextDecoration = @(),
        [ValidateSet('uppercase', 'lowercase', 'capitalize')][string[]] $TextTransform = @(),
        [ValidateSet('rtl')][string[]] $Direction = @(),
        [switch] $LineBreak
    )

    $newHTMLTextSplat = @{
        Alignment       = $Alignment
        FontSize        = $FontSize
        TextTransform   = $TextTransform
        Text            = $Text
        Color           = $Color
        FontFamily      = $FontFamily
        Direction       = $Direction
        FontStyle       = $FontStyle
        TextDecoration  = $TextDecoration
        BackGroundColor = $BackGroundColor
        FontVariant     = $FontVariant
        FontWeight      = $FontWeight
        LineBreak       = $LineBreak
    }

    New-HTMLListItem @newHTMLTextSplat
}
Register-ArgumentCompleter -CommandName EmailListItem -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName EmailListItem -ParameterName BackgroundColor -ScriptBlock { $Script:RGBColors.Keys }
function EmailOptions {
    [CmdletBinding()]
    param(
        [ValidateSet('Low', 'Normal', 'High')] [string] $Priority = 'Normal',
        [ValidateSet('None', 'OnSuccess', 'OnFailure', 'Delay', 'Never')] $DeliveryNotifications = 'None',
        [string] $Encoding = 'Unicode'
    )

    [PsCustomObject] @{
        Type                  = 'HeaderOptions'
        Encoding              = $Encoding
        DeliveryNotifications = $DeliveryNotifications
        Priority              = $Priority
    }
}
function EmailReplyTo {
    [CmdletBinding()]
    param(
        [string] $Address
    )

    [PsCustomObject] @{
        Type    = 'HeaderReplyTo'
        Address = $Address
    }
}
function EmailServer {
    [CmdletBinding()]
    param(
        [string] $Server,
        [int] $Port = 587,
        [string] $UserName,
        [string] $Password,
        [switch] $PasswordAsSecure,
        [switch] $PasswordFromFile,
        [switch] $SSL
    )

    [PsCustomObject] @{
        Type             = 'HeaderServer'
        Server           = $Server
        Port             = $Port
        UserName         = $UserName
        Password         = $Password
        PasswordAsSecure = $PasswordAsSecure
        PasswordFromFile = $PasswordFromFile
        SSL              = $SSL
    }
}
function EmailSubject {
    [CmdletBinding()]
    param(
        [string] $Subject
    )

    [PsCustomObject] @{
        Type    = 'HeaderSubject'
        Subject = $Subject
    }
}
function EmailText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)][ScriptBlock] $TextBlock,
        [string[]] $Text,
        [string[]] $Color = @(),
        [string[]] $BackGroundColor = @(),
        [alias('Size')][int[]] $FontSize = @(),
        [ValidateSet('normal', 'bold', 'bolder', 'lighter', '100', '200', '300', '400', '500', '600', '700', '800', '900')][string[]] $FontWeight = @(),
        [ValidateSet('normal', 'italic', 'oblique')][string[]] $FontStyle = @(),
        [ValidateSet('none', 'line-through', 'overline', 'underline')][string[]] $TextDecoration = @(),
        [ValidateSet('normal', 'small-caps')][string[]] $FontVariant = @(),
        [string[]] $FontFamily = @(),
        [ValidateSet('left', 'center', 'right', 'justify')][string[]] $Alignment = @(),
        [ValidateSet('uppercase', 'lowercase', 'capitalize')][string[]] $TextTransform = @(),
        [ValidateSet('rtl')][string[]] $Direction = @(),
        [switch] $LineBreak,
        [switch] $SkipParagraph
    )
    if ($TextBlock) {
        $Text = (Invoke-Command -ScriptBlock $TextBlock)
        #if ($Text.Count) {
        #    $LineBreak = $false
        #}
    }

    $newHTMLTextSplat = @{
        Alignment       = $Alignment
        FontSize        = $FontSize
        TextTransform   = $TextTransform
        Text            = $Text
        Color           = $Color
        FontFamily      = $FontFamily
        Direction       = $Direction
        FontStyle       = $FontStyle
        TextDecoration  = $TextDecoration
        BackGroundColor = $BackGroundColor
        FontVariant     = $FontVariant
        FontWeight      = $FontWeight
        LineBreak       = $LineBreak
        SkipParagraph   = $SkipParagraph
    }

    New-HTMLText @newHTMLTextSplat
}
Register-ArgumentCompleter -CommandName EmailText -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName EmailText -ParameterName BackgroundColor -ScriptBlock { $Script:RGBColors.Keys }
function EmailTextBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)][ScriptBlock] $TextBlock,
        [string[]] $Color = @(),
        [string[]] $BackGroundColor = @(),
        [alias('Size')][int[]] $FontSize = @(),
        [ValidateSet('normal', 'bold', 'bolder', 'lighter', '100', '200', '300', '400', '500', '600', '700', '800', '900')][string[]] $FontWeight = @(),
        [ValidateSet('normal', 'italic', 'oblique')][string[]] $FontStyle = @(),
        [ValidateSet('none', 'line-through', 'overline', 'underline')][string[]] $TextDecoration = @(),
        [ValidateSet('normal', 'small-caps')][string[]] $FontVariant = @(),
        [string[]] $FontFamily = @(),
        [ValidateSet('left', 'center', 'right', 'justify')][string[]] $Alignment = @(),
        [ValidateSet('uppercase', 'lowercase', 'capitalize')][string[]] $TextTransform = @(),
        [ValidateSet('rtl')][string[]] $Direction = @(),
        [switch] $LineBreak
    )
    if ($TextBlock) {
        $Text = (Invoke-Command -ScriptBlock $TextBlock)
        if ($Text.Count) {
            $LineBreak = $true
        }
    }
    foreach ($T in $Text) {
        $newHTMLTextSplat = @{
            Alignment       = $Alignment
            FontSize        = $FontSize
            TextTransform   = $TextTransform
            Text            = $T
            Color           = $Color
            FontFamily      = $FontFamily
            Direction       = $Direction
            FontStyle       = $FontStyle
            TextDecoration  = $TextDecoration
            BackGroundColor = $BackGroundColor
            FontVariant     = $FontVariant
            FontWeight      = $FontWeight
            LineBreak       = $LineBreak
        }
        New-HTMLText @newHTMLTextSplat -SkipParagraph
    }
}
Register-ArgumentCompleter -CommandName EmailTextBox -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName EmailTextBox -ParameterName BackgroundColor -ScriptBlock { $Script:RGBColors.Keys }

function EmailTo {
    [CmdletBinding()]
    param(
        [string[]] $Addresses
    )

    [PsCustomObject] @{
        Type      = 'HeaderTo'
        Addresses = $Addresses
    }
}
function New-CalendarEvent {
    [alias('CalendarEvent')]
    [CmdletBinding()]
    param(
        [string] $Title,
        [string] $Description,
        [DateTime] $StartDate,
        [nullable[DateTime]] $EndDate,
        [string] $Constraint,
        [string] $Color
    )

    $Object = [PSCustomObject] @{
        Type     = 'CalendarEvent'
        Settings = [ordered] @{
            title       = $Title
            description = $Description
            constraint  = $Constraint
            #      url: 'http://google.com/',
            color       = ConvertFrom-Color -Color $Color
        }
    }
    if ($StartDate) {
        $Object.Settings.start = Get-Date -Date ($StartDate) -Format "yyyy-MM-ddTHH:mm:ss"
    }
    if ($EndDate) {
        $Object.Settings.end = Get-Date -Date ($EndDate) -Format "yyyy-MM-ddTHH:mm:ss"
    }

    Remove-EmptyValues -Hashtable $Object.Settings -Recursive #-Rerun 2
    $Object
}
Register-ArgumentCompleter -CommandName New-CalendarEvent -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
<#
    events: [
    {
        title: 'rrule event',
        rrule: {
        dtstart: '2019-08-09T13:00:00',
        // until: '2019-08-01',
        freq: 'weekly'
        },
        duration: '02:00'
    }
    ],
events: [{
        title: 'Business Lunch',
        start: '2019-08-03T13:00:00',
        constraint: 'businessHours'
    },
    {
        title: 'Meeting',
        start: '2019-08-13T11:00:00',
        constraint: 'availableForMeeting', // defined below
        color: '#257e4a'
    },
    {
        title: 'Conference',
        start: '2019-08-18',
        end: '2019-08-20'
    },
    {
        title: 'Party',
        start: '2019-08-29T20:00:00'
    },

    // areas where "Meeting" must be dropped
    {
        groupId: 'availableForMeeting',
        start: '2019-08-11T10:00:00',
        end: '2019-08-11T16:00:00',
        rendering: 'background'
    },
    {
        groupId: 'availableForMeeting',
        start: '2019-08-13T10:00:00',
        end: '2019-08-13T16:00:00',
        rendering: 'background'
    },

    // red areas where no events can be dropped
    {
        start: '2019-08-24',
        end: '2019-08-28',
        overlap: false,
        rendering: 'background',
        color: '#ff9f89'
    },
    {
        start: '2019-08-06',
        end: '2019-08-08',
        overlap: false,
        rendering: 'background',
        color: '#ff9f89'
    }
]
#>

function New-ChartAxisX {
    [alias('ChartCategory', 'ChartAxisX', 'New-ChartCategory')]
    [CmdletBinding()]
    param(
        [alias('Name')][Array] $Names,
        [alias('Title')][string] $TitleText,
        [ValidateSet('datetime', 'category', 'numeric')][string] $Type = 'category',
        [int] $MinValue,
        [int] $MaxValue
        #[ValidateSet('top', 'topRight', 'left', 'right', 'bottom', '')][string] $LegendPosition = '',
        # [string[]] $Color
    )
    [PSCustomObject] @{
        ObjectType = 'ChartAxisX'
        ChartAxisX = @{
            Names     = $Names
            Type      = $Type
            TitleText = $TitleText
            Min       = $MinValue
            Max       = $MaxValue
        }

        #   LegendPosition = $LegendPosition
        #   Color          = $Color
    }

    # https://apexcharts.com/docs/options/xaxis/
}

<# We can build this:
   xaxis: {
        type: 'category',
        categories: [],
        labels: {
            show: true,
            rotate: -45,
            rotateAlways: false,
            hideOverlappingLabels: true,
            showDuplicates: false,
            trim: true,
            minHeight: undefined,
            maxHeight: 120,
            style: {
                colors: [],
                fontSize: '12px',
                fontFamily: 'Helvetica, Arial, sans-serif',
                cssClass: 'apexcharts-xaxis-label',
            },
            offsetX: 0,
            offsetY: 0,
            format: undefined,
            formatter: undefined,
            datetimeFormatter: {
                year: 'yyyy',
                month: "MMM 'yy",
                day: 'dd MMM',
                hour: 'HH:mm',
            },
        },
        axisBorder: {
            show: true,
            color: '#78909C',
            height: 1,
            width: '100%',
            offsetX: 0,
            offsetY: 0
        },
        axisTicks: {
            show: true,
            borderType: 'solid',
            color: '#78909C',
            height: 6,
            offsetX: 0,
            offsetY: 0
        },
        tickAmount: undefined,
        tickPlacement: 'between',
        min: undefined,
        max: undefined,
        range: undefined,
        floating: false,
        position: 'bottom',
        title: {
            text: undefined,
            offsetX: 0,
            offsetY: 0,
            style: {
                color: undefined,
                fontSize: '12px',
                fontFamily: 'Helvetica, Arial, sans-serif',
                cssClass: 'apexcharts-xaxis-title',
            },
        },
        crosshairs: {
            show: true,
            width: 1,
            position: 'back',
            opacity: 0.9,
            stroke: {
                color: '#b6b6b6',
                width: 0,
                dashArray: 0,
            },
            fill: {
                type: 'solid',
                color: '#B1B9C4',
                gradient: {
                    colorFrom: '#D8E3F0',
                    colorTo: '#BED1E6',
                    stops: [0, 100],
                    opacityFrom: 0.4,
                    opacityTo: 0.5,
                },
            },
            dropShadow: {
                enabled: false,
                top: 0,
                left: 0,
                blur: 1,
                opacity: 0.4,
            },
        },
        tooltip: {
            enabled: true,
            formatter: undefined,
            offsetY: 0,
        },
    }

#>
function New-ChartAxisY {
    [alias('ChartAxisY')]
    [CmdletBinding()]
    param(
        [switch] $Show,
        [switch] $ShowAlways,
        [string] $TitleText,
        [ValidateSet('90', '270')][string] $TitleRotate = '90',
        [int] $TitleOffsetX = 0,
        [int] $TitleOffsetY = 0,
        [string] $TitleStyleColor,
        [int] $TitleStyleFontSize = 12,
        [string] $TitleStylefontFamily = 'Helvetica, Arial, sans-serif',
        [int] $MinValue,
        [int] $MaxValue
        #[ValidateSet('top', 'topRight', 'left', 'right', 'bottom', '')][string] $LegendPosition = '',
        # [string[]] $Color
    )
    [PSCustomObject] @{
        ObjectType = 'ChartAxisY'
        ChartAxisY = @{
            Show                 = $Show.IsPresent
            ShowAlways           = $ShowAlways.IsPresent
            TitleText            = $TitleText
            TitleRotate          = $TitleRotate
            TitleOffsetX         = $TitleOffsetX
            TitleOffsetY         = $TitleOffsetY
            TitleStyleColor      = $TitleStyleColor
            TitleStyleFontSize   = $TitleStyleFontSize
            TitleStylefontFamily = $TitleStylefontFamily
            Min                  = $MinValue
            Max                  = $MaxValue
        }
    }

    # https://apexcharts.com/docs/options/yaxis/
}
Register-ArgumentCompleter -CommandName New-ChartAxisY -ParameterName TitleStyleColor -ScriptBlock { $Script:RGBColors.Keys }

<# We can build this
    yaxis: {
        show: true,
        showAlways: true,
        seriesName: undefined,
        opposite: false,
        reversed: false,
        logarithmic: false,
        tickAmount: 6,
        min: 6,
        max: 6,
        forceNiceScale: false,
        floating: false,
        decimalsInFloat: undefined,
        labels: {
            show: true,
            align: 'right',
            minWidth: 0,
            maxWidth: 160,
            style: {
                color: undefined,
                fontSize: '12px',
                fontFamily: 'Helvetica, Arial, sans-serif',
                cssClass: 'apexcharts-yaxis-label',
            },
            offsetX: 0,
            offsetY: 0,
            rotate: 0,
            formatter: (value) => { return val },
        },
        axisBorder: {
            show: true,
            color: '#78909C',
            offsetX: 0,
            offsetY: 0
        },
        axisTicks: {
            show: true,
            borderType: 'solid',
            color: '#78909C',
            width: 6,
            offsetX: 0,
            offsetY: 0
        },
        title: {
            text: undefined,
            rotate: -90,
            offsetX: 0,
            offsetY: 0,
            style: {
                color: undefined,
                fontSize: '12px',
                fontFamily: 'Helvetica, Arial, sans-serif',
                cssClass: 'apexcharts-yaxis-title',
            },
        },
        crosshairs: {
            show: true,
            position: 'back',
            stroke: {
                color: '#b6b6b6',
                width: 1,
                dashArray: 0,
            },
        },
        tooltip: {
            enabled: true,
            offsetX: 0,
        },

    }

#>
function New-ChartBar {
    [alias('ChartBar')]
    [CmdletBinding()]
    param(
        [string] $Name,
        [object] $Value
    )
    [PSCustomObject] @{
        ObjectType = 'Bar'
        Name       = $Name
        Value      = $Value
    }
}
function New-ChartBarOptions {
    [alias('ChartBarOptions')]
    [CmdletBinding()]
    param(
        [ValidateSet('bar', 'barStacked', 'barStacked100Percent')] $Type = 'bar',
        [bool] $DataLabelsEnabled = $true,
        [int] $DataLabelsOffsetX = -6,
        [string] $DataLabelsFontSize = '12px',
        [string] $DataLabelsColor,
        [alias('PatternedColors')][switch] $Patterned,
        [alias('GradientColors')][switch] $Gradient,
        [switch] $Distributed,
        [switch] $Vertical

    )

    if ($null -ne $PSBoundParameters.Patterned) {
        $PatternedColors = $Patterned.IsPresent
    } else {
        $PatternedColors = $null
    }
    if ($null -ne $PSBoundParameters.Gradient) {
        $GradientColors = $Gradient.IsPresent
    } else {
        $GradientColors = $null
    }

    [PSCustomObject] @{
        ObjectType         = 'BarOptions'
        Type               = $Type
        Title              = $Title
        TitleAlignment     = $TitleAlignment
        Horizontal         = -not $Vertical.IsPresent
        DataLabelsEnabled  = $DataLabelsEnabled
        DataLabelsOffsetX  = $DataLabelsOffsetX
        DataLabelsFontSize = $DataLabelsFontSize
        DataLabelsColor    = $DataLabelsColor
        PatternedColors    = $PatternedColors
        GradientColors     = $GradientColors
        Distributed        = $Distributed.IsPresent
    }
}

Register-ArgumentCompleter -CommandName New-ChartBarOptions -ParameterName LineColor -ScriptBlock { $Script:RGBColors.Keys }
function New-ChartDonut {
    [alias('ChartDonut')]
    [CmdletBinding()]
    param(
        [string] $Name,
        [object] $Value,
        [string] $Color
    )

    [PSCustomObject] @{
        ObjectType = 'Donut'
        Name       = $Name
        Value      = $Value
        Color      = $Color
    }
}
Register-ArgumentCompleter -CommandName New-ChartDonut -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
function New-ChartGrid {
    [alias('ChartGrid')]
    [CmdletBinding()]
    param(
        [switch] $Show,
        [string] $BorderColor,
        [int] $StrokeDash, #: 0,
        [ValidateSet('front', 'back', 'default')][string] $Position = 'default',
        [switch] $xAxisLinesShow,
        [switch] $yAxisLinesShow,
        [string[]] $RowColors,
        [double] $RowOpacity = 0.5, # valid range 0 - 1
        [string[]] $ColumnColors,
        [double] $ColumnOpacity = 0.5, # valid range 0 - 1
        [int] $PaddingTop,
        [int] $PaddingRight,
        [int] $PaddingBottom,
        [int] $PaddingLeft
    )
    [PSCustomObject] @{
        ObjectType = 'ChartGrid'
        Grid       = @{
            Show           = $Show.IsPresent
            BorderColor    = $BorderColor
            StrokeDash     = $StrokeDash
            Position       = $Position
            xAxisLinesShow = $xAxisLinesShow.IsPresent
            yAxisLinesShow = $yAxisLinesShow.IsPresent
            RowColors      = $RowColors
            RowOpacity     = $RowOpacity
            ColumnColors   = $ColumnColors
            ColumnOpacity  = $ColumnOpacity
            PaddingTop     = $PaddingTop
            PaddingRight   = $PaddingRight
            PaddingBottom  = $PaddingBottom
            PaddingLeft    = $PaddingLeft
        }
    }
    # https://apexcharts.com/docs/options/xaxis/
}
Register-ArgumentCompleter -CommandName New-ChartGrid -ParameterName BorderColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-ChartGrid -ParameterName RowColors -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-ChartGrid -ParameterName ColumnColors -ScriptBlock { $Script:RGBColors.Keys }
function New-ChartLegend {
    [alias('ChartLegend')]
    [CmdletBinding()]
    param(
        [Array] $Names,
        [ValidateSet('top', 'topRight', 'left', 'right', 'bottom', 'default')][string] $LegendPosition = 'default',
        [string[]] $Color
    )

    #$Colors = "Red","Blue","orange"
    #foreach ($_ in $Color) {
    #    $Colors.Add($Color)
    #}
    [PSCustomObject] @{
        ObjectType     = 'Legend'
        Names          = $Names
        LegendPosition = $LegendPosition
        Color          = $Color
    }
}
Register-ArgumentCompleter -CommandName New-ChartLegend -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
function New-ChartLine {
    [alias('ChartLine')]
    [CmdletBinding()]
    param(
        [string] $Name,
        [object] $Value,
        [string] $Color,
        [ValidateSet('straight', 'smooth', 'stepline')] $Curve = 'straight',
        [int] $Width = 6,
        [ValidateSet('butt', 'square', 'round')][string] $Cap = 'butt',
        [int] $Dash = 0
    )
    [PSCustomObject] @{
        ObjectType = 'Line'
        Name       = $Name
        Value      = $Value
        LineColor  = $Color
        LineCurve  = $Curve
        LineWidth  = $Width
        LineCap    = $Cap
        LineDash   = $Dash
    }
}

Register-ArgumentCompleter -CommandName New-ChartLine -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
function New-ChartPie {
    [alias('ChartPie')]
    [CmdletBinding()]
    param(
        [string] $Name,
        [object] $Value,
        [string] $Color
    )

    [PSCustomObject] @{
        ObjectType = 'Pie'
        Name       = $Name
        Value      = $Value
        Color      = $Color
    }
}

Register-ArgumentCompleter -CommandName New-ChartPie -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
function New-ChartRadial {
    [alias('ChartRadial')]
    [CmdletBinding()]
    param(
        [string] $Name,
        [object] $Value,
        [string] $Color
    )

    [PSCustomObject] @{
        ObjectType = 'Radial'
        Name       = $Name
        Value      = $Value
        Color      = $Color
    }
}

Register-ArgumentCompleter -CommandName New-ChartRadial -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
function New-ChartTheme {
    [alias('ChartTheme')]
    [CmdletBinding()]
    param(
        [ValidateSet('light', 'dark')][string] $Mode = 'light',
        [ValidateSet(
            'palette1',
            'palette2',
            'palette3',
            'palette4',
            'palette5',
            'palette6',
            'palette7',
            'palette8',
            'palette9',
            'palette10'
        )
        ][string] $Palette = 'palette1',
        [switch] $Monochrome,
        [string] $Color = "DodgerBlue",
        [ValidateSet('light', 'dark')][string] $ShadeTo = 'light',
        [double] $ShadeIntensity = 0.65
    )

    [PSCustomObject] @{
        ObjectType = 'Theme'
        Theme      = @{
            Mode           = $Mode
            Palette        = $Palette
            Monochrome     = $Monochrome.IsPresent
            Color          = $Color
            ShadeTo        = $ShadeTo
            ShadeIntensity = $ShadeIntensity
        }
    }
}

Register-ArgumentCompleter -CommandName New-ChartTheme -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
function New-ChartToolbar {
    [alias('ChartToolbar')]
    [CmdletBinding()]
    param(
        [switch] $Download,
        [switch] $Selection,
        [switch] $Zoom,
        [switch] $ZoomIn,
        [switch] $ZoomOut,
        [switch] $Pan,
        [switch] $Reset,
        [ValidateSet('zoom', 'selection', 'pan')][string] $AutoSelected = 'zoom'
    )

    [PSCustomObject] @{
        ObjectType = 'Toolbar'
        Toolbar    = @{
            #Show         = $Show.IsPresent
            #tools        = [ordered] @{
            download     = $Download.IsPresent
            selection    = $Selection.IsPresent
            zoom         = $Zoom.IsPresent
            zoomin       = $ZoomIn.IsPresent
            zoomout      = $ZoomOut.IsPresent
            pan          = $Pan.IsPresent
            reset        = $Reset.IsPresent
            #}
            autoSelected = $AutoSelected
        }
    }
}
function New-DiagramEvent {
    [CmdletBinding()]
    param(
        #[switch] $FadeSearch,
        [string] $ID,
        [nullable[int]] $ColumnID
    )

    $Object = [PSCustomObject] @{
        Type     = 'DiagramEvent'
        Settings = @{
            # OnClick = $OnClick.IsPresent
            ID       = $ID
            # FadeSearch = $FadeSearch.IsPresent
            ColumnID = $ColumnID
        }
    }
    $Object
}
function New-DiagramLink {
    [alias('DiagramEdge', 'DiagramEdges', 'New-DiagramEdge', 'DiagramLink')]
    [CmdletBinding()]
    param(
        [string[]] $From,
        [string[]] $To,
        [string] $Label,
        [nullable[bool]] $ArrowsToEnabled,
        [nullable[int]] $ArrowsToScaleFacto,
        [ValidateSet('arrow', 'bar', 'circle')][string] $ArrowsToType,
        [nullable[bool]] $ArrowsMiddleEnabled,
        [nullable[int]]$ArrowsMiddleScaleFactor,
        [ValidateSet('arrow', 'bar', 'circle')][string] $ArrowsMiddleType,
        [nullable[bool]] $ArrowsFromEnabled,
        [nullable[int]] $ArrowsFromScaleFactor,
        [ValidateSet('arrow', 'bar', 'circle')][string] $ArrowsFromType,
        [nullable[bool]]$ArrowStrikethrough,
        [nullable[bool]] $Chosen,
        [string] $Color,
        [string] $ColorHighlight,
        [string] $ColorHover,
        [ValidateSet('true', 'false', 'from', 'to', 'both')][string]$ColorInherit,
        [nullable[double]] $ColorOpacity, # range between 0 and 1
        [nullable[bool]] $Dashes,
        [string] $Length,
        [string] $FontColor,
        [nullable[int]] $FontSize, #// px
        [string] $FontName,
        [string] $FontBackground,
        [nullable[int]] $FontStrokeWidth, #// px
        [string] $FontStrokeColor,
        [ValidateSet('center', 'left')][string] $FontAlign,
        [ValidateSet('false', 'true', 'markdown', 'html')][string]$FontMulti,
        [nullable[int]] $FontVAdjust,
        [nullable[int]] $WidthConstraint
    )
    $Object = [PSCustomObject] @{
        Type     = 'DiagramLink'
        Settings = @{
            from = $From
            to   = $To
        }
        Edges    = @{
            label              = $Label
            length             = $Length
            arrows             = [ordered]@{
                to     = [ordered]@{
                    enabled     = $ArrowsToEnabled
                    scaleFactor = $ArrowsToScaleFactor
                    type        = $ArrowsToType
                }
                middle = [ordered]@{
                    enabled     = $ArrowsMiddleEnabled
                    scaleFactor = $ArrowsMiddleScaleFactor
                    type        = $ArrowsMiddleType
                }
                from   = [ordered]@{
                    enabled     = $ArrowsFromEnabled
                    scaleFactor = $ArrowsFromScaleFactor
                    type        = $ArrowsFromType
                }
            }
            arrowStrikethrough = $ArrowStrikethrough
            chosen             = $Chosen
            color              = [ordered]@{
                color     = ConvertFrom-Color -Color $Color
                highlight = ConvertFrom-Color -Color $ColorHighlight
                hover     = ConvertFrom-Color -Color $ColorHover
                inherit   = $ColorInherit
                opacity   = $ColorOpacity
            }
            font               = [ordered]@{
                color       = ConvertFrom-Color -Color $FontColor
                size        = $FontSize
                face        = $FontName
                background  = ConvertFrom-Color -Color $FontBackground
                strokeWidth = $FontStrokeWidth
                strokeColor = ConvertFrom-Color -Color $FontStrokeColor
                align       = $FontAlign
                multi       = $FontMulti
                vadjust     = $FontVAdjust
            }
            dashes             = $Dashes
            widthConstraint    = $WidthConstraint
        }
    }
    Remove-EmptyValues -Hashtable $Object.Settings -Recursive
    Remove-EmptyValues -Hashtable $Object.Edges -Recursive -Rerun 2
    $Object
}
Register-ArgumentCompleter -CommandName New-DiagramLink -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramLink -ParameterName ColorHighlight -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramLink -ParameterName ColorHover -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramLink -ParameterName FontColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramLink -ParameterName FontBackground -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramLink -ParameterName FontStrokeColor -ScriptBlock { $Script:RGBColors.Keys }

<#
  // these are all options in full.
  var options = {
    edges:{
      arrows: {
        to:     {enabled: false, scaleFactor:1, type:'arrow'},
        middle: {enabled: false, scaleFactor:1, type:'arrow'},
        from:   {enabled: false, scaleFactor:1, type:'arrow'}
      },
      arrowStrikethrough: true,
      chosen: true,
      color: {
        color:'#848484',
        highlight:'#848484',
        hover: '#848484',
        inherit: 'from',
        opacity:1.0
      },
      dashes: false,
      font: {
        color: '#343434',
        size: 14, // px
        face: 'arial',
        background: 'none',
        strokeWidth: 2, // px
        strokeColor: '#ffffff',
        align: 'horizontal',
        multi: false,
        vadjust: 0,
        bold: {
          color: '#343434',
          size: 14, // px
          face: 'arial',
          vadjust: 0,
          mod: 'bold'
        },
        ital: {
          color: '#343434',
          size: 14, // px
          face: 'arial',
          vadjust: 0,
          mod: 'italic',
        },
        boldital: {
          color: '#343434',
          size: 14, // px
          face: 'arial',
          vadjust: 0,
          mod: 'bold italic'
        },
        mono: {
          color: '#343434',
          size: 15, // px
          face: 'courier new',
          vadjust: 2,
          mod: ''
        }
      },
      hidden: false,
      hoverWidth: 1.5,
      label: undefined,
      labelHighlightBold: true,
      length: undefined,
      physics: true,
      scaling:{
        min: 1,
        max: 15,
        label: {
          enabled: true,
          min: 14,
          max: 30,
          maxVisible: 30,
          drawThreshold: 5
        },
        customScalingFunction: function (min,max,total,value) {
          if (max === min) {
            return 0.5;
          }
          else {
            var scale = 1 / (max - min);
            return Math.max(0,(value - min)*scale);
          }
        }
      },
      selectionWidth: 1,
      selfReferenceSize:20,
      shadow:{
        enabled: false,
        color: 'rgba(0,0,0,0.5)',
        size:10,
        x:5,
        y:5
      },
      smooth: {
        enabled: true,
        type: "dynamic",
        roundness: 0.5
      },
      title:undefined,
      value: undefined,
      width: 1,
      widthConstraint: false
    }
  }

  network.setOptions(options);
  #>
function New-DiagramNode {
    [alias('DiagramNode')]
    [CmdLetBinding(DefaultParameterSetName = 'Shape')]
    param(
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][ScriptBlock] $TextBox,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][string] $Id,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")] [string] $Label,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][string[]] $To,
        [parameter(ParameterSetName = "Shape")][string][ValidateSet(
            'circle', 'dot', 'diamond', 'ellipse', 'database', 'box', 'square', 'triangle', 'triangleDown', 'text', 'star', 'hexagon')] $Shape,
        [parameter(ParameterSetName = "Image")][ValidateSet('squareImage', 'circularImage')][string] $ImageType,
        [parameter(ParameterSetName = "Image")][uri] $Image,
        #[string] $BrokenImage,
        #[string] $ImagePadding,
        #[string] $ImagePaddingLeft,
        #[string] $ImagePaddingRight,
        #[string] $ImagePaddingTop,
        #[string] $ImagePaddingBottom,
        #[string] $UseImageSize,
        #[alias('BackgroundColor')][string] $Color,
        #[string] $Border,
        #[string] $HighlightBackground,
        #[string] $HighlightBorder,
        #[string] $HoverBackground,
        #[string] $HoverBorder,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][nullable[int]] $BorderWidth,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][nullable[int]] $BorderWidthSelected,
        [parameter(ParameterSetName = "Image")][string] $BrokenImages,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][nullable[bool]] $Chosen,
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][string] $ColorBorder,
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][string] $ColorBackground,
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][string] $ColorHighlightBorder,
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][string] $ColorHighlightBackground,
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][string] $ColorHoverBorder,
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][string] $ColorHoverBackground,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][nullable[bool]]$FixedX,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][nullable[bool]]$FixedY,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][string] $FontColor,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][nullable[int]] $FontSize, #// px
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][string] $FontName,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][string] $FontBackground,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][nullable[int]] $FontStrokeWidth, #// px
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][string] $FontStrokeColor,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][ValidateSet('center', 'left')][string] $FontAlign,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][ValidateSet('false', 'true', 'markdown', 'html')][string]$FontMulti,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][nullable[int]] $FontVAdjust,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][nullable[int]] $Size,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][nullable[int]] $X,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][nullable[int]] $Y,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")][switch] $IconAsImage,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")][string] $IconColor,
        # ICON BRANDS
        [ArgumentCompleter(
            {
                param($CommandName, $ParameterName, $WordToComplete, $CommandAst, $FakeBoundParameters)
                ($Global:HTMLIcons.FontAwesomeBrands.Keys)
            }
        )]
        [ValidateScript(
            {
                $_ -in (($Global:HTMLIcons.FontAwesomeBrands.Keys))
            }
        )]
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [string] $IconBrands,

        # ICON REGULAR
        [ArgumentCompleter(
            {
                param($CommandName, $ParameterName, $WordToComplete, $CommandAst, $FakeBoundParameters)
                ($Global:HTMLIcons.FontAwesomeRegular.Keys)
            }
        )]
        [ValidateScript(
            {
                $_ -in (($Global:HTMLIcons.FontAwesomeRegular.Keys))
            }
        )]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [string] $IconRegular,

        # ICON SOLID
        [ArgumentCompleter(
            {
                param($CommandName, $ParameterName, $WordToComplete, $CommandAst, $FakeBoundParameters)
                ($Global:HTMLIcons.FontAwesomeSolid.Keys)
            }
        )]
        [ValidateScript(
            {
                $_ -in (($Global:HTMLIcons.FontAwesomeSolid.Keys))
            }
        )]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [string] $IconSolid,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][nullable[int]] $Level,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][nullable[int]] $HeightConstraintMinimum,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][ValidateSet('top', 'middle', 'bottom')][string] $HeightConstraintVAlign,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][nullable[int]] $WidthConstraintMinimum,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [parameter(ParameterSetName = "Image")]
        [parameter(ParameterSetName = "Shape")][nullable[int]] $WidthConstraintMaximum
    )

    if (-not $Label) {
        Write-Warning 'New-DiagramNode - Label is required. Skipping node.'
        return
    }

    $Object = [PSCustomObject] @{
        Type     = 'DiagramNode'
        Settings = @{ }
        Edges    = @{ }
    }
    $Icon = @{ } # Reset value, just in case

    # If ID is not defined use label
    if (-not $ID) {
        $ID = $Label
    }

    if ($IconBrands -or $IconRegular -or $IconSolid) {
        if ($IconBrands) {
            if (-not $IconAsImage) {
                # Workaround using image for Fonts
                # https://use.fontawesome.com/releases/v5.11.2/svgs/brands/accessible-icon.svg
                <# Until all Icons work, using images instead. Currently only Brands work fine / Solid/Regular is weird #>
                $NodeShape = 'icon'
                $icon = @{
                    face   = '"Font Awesome 5 Brands"'
                    code   = -join ('\u', $Global:HTMLIcons.FontAwesomeBrands[$IconBrands])    # "\uf007"
                    color  = ConvertFrom-Color -Color $IconColor
                    weight = 'bold'
                }

            } else {
                $NodeShape = 'image'
                $Image = -join ($Script:Configuration.Features.FontsAwesome.Other.Link, 'brands/', $IconBrands, '.svg')
            }
        } elseif ($IconRegular) {
            if (-not $IconAsImage) {
                $NodeShape = 'icon'
                $icon = @{
                    face   = '"Font Awesome 5 Free"'
                    code   = -join ('\u', $Global:HTMLIcons.FontAwesomeRegular[$IconRegular])    # "\uf007"
                    color  = ConvertFrom-Color -Color $IconColor
                    weight = 'bold'
                }
            } else {
                $NodeShape = 'image'
                $Image = -join ($Script:Configuration.Features.FontsAwesome.Other.Link, 'regular/', $IconRegular, '.svg')
            }
        } else {
            if (-not $IconAsImage) {
                $NodeShape = 'icon'
                $icon = @{
                    face   = '"Font Awesome 5 Free"'
                    code   = -join ('\u', $Global:HTMLIcons.FontAwesomeSolid[$IconSolid])    # "\uf007"
                    color  = ConvertFrom-Color -Color $IconColor
                    weight = 'bold'
                }

            } else {
                $NodeShape = 'image'
                $Image = -join ($Script:Configuration.Features.FontsAwesome.Other.Link, 'solid/', $IconSolid, '.svg')
            }
        }
    } elseif ($Image) {
        if ($ImageType -eq 'squareImage') {
            $NodeShape = 'image'
        } else {
            $NodeShape = 'circularImage'
        }
    } else {
        $NodeShape = $Shape
    }

    if ($To) {
        $Object.Edges = @{
            from = if ($To) { $Id } else { '' }
            to   = if ($To) { $To } else { '' }
        }
    }
    $Object.Settings = [ordered] @{
        id                  = $Id
        label               = $Label
        shape               = $NodeShape

        image               = $Image
        icon                = $icon

        level               = $Level


        borderWidth         = $BorderWidth
        borderWidthSelected = $BorderWidthSelected
        brokenImage         = $BrokenImage

        chosen              = $Chosen
        color               = [ordered]@{
            border     = ConvertFrom-Color -Color $ColorBorder
            background = ConvertFrom-Color -Color $ColorBackground
            highlight  = [ordered]@{
                border     = ConvertFrom-Color -Color $ColorHighlightBorder
                background = ConvertFrom-Color -Color $ColorHighlightBackground
            }
            hover      = [ordered]@{
                border     = ConvertFrom-Color -Color $ColorHoverBorder
                background = ConvertFrom-Color -Color $ColorHoverBackground
            }
        }
        fixed               = [ordered]@{
            x = $FixedX
            y = $FixedY
        }
        font                = [ordered]@{
            color       = ConvertFrom-Color -Color $FontColor
            size        = $FontSize
            face        = $FontName
            background  = ConvertFrom-Color -Color $FontBackground
            strokeWidth = $FontStrokeWidth
            strokeColor = ConvertFrom-Color -Color $FontStrokeColor
            align       = $FontAlign
            multi       = $FontMulti
            vadjust     = $FontVAdjust
        }
        size                = $Size
        heightConstraint    = @{
            minimum = $HeightConstraintMinimum
            valign  = $HeightConstraintVAlign
        }
        widthConstraint     = @{
            minimum = $WidthConstraintMinimum
            maximum = $WidthConstraintMaximum
        }
        x                   = $X
        y                   = $Y
    }

    Remove-EmptyValues -Hashtable $Object.Settings -Recursive -Rerun 2
    Remove-EmptyValues -Hashtable $Object.Edges -Recursive
    $Object
}
Register-ArgumentCompleter -CommandName New-DiagramNode -ParameterName ColorBorder -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramNode -ParameterName ColorBackground -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramNode -ParameterName ColorHighlightBorder -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramNode -ParameterName ColorHighlightBackground -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramNode -ParameterName ColorHoverBorder -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramNode -ParameterName ColorHoverBackground -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramNode -ParameterName FontColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramNode -ParameterName FontBackground -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramNode -ParameterName FontStrokeColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramNode -ParameterName IconColor -ScriptBlock { $Script:RGBColors.Keys }
<#
// these are all options in full.
var options = {
  nodes:{
    borderWidth: 1,
    borderWidthSelected: 2,
    brokenImage:undefined,
    chosen: true,
    color: {
      border: '#2B7CE9',
      background: '#97C2FC',
      highlight: {
        border: '#2B7CE9',
        background: '#D2E5FF'
      },
      hover: {
        border: '#2B7CE9',
        background: '#D2E5FF'
      }
    },
    fixed: {
      x:false,
      y:false
    },
    font: {
      color: '#343434',
      size: 14, // px
      face: 'arial',
      background: 'none',
      strokeWidth: 0, // px
      strokeColor: '#ffffff',
      align: 'center',
      multi: false,
      vadjust: 0,
      bold: {
        color: '#343434',
        size: 14, // px
        face: 'arial',
        vadjust: 0,
        mod: 'bold'
      },
      ital: {
        color: '#343434',
        size: 14, // px
        face: 'arial',
        vadjust: 0,
        mod: 'italic',
      },
      boldital: {
        color: '#343434',
        size: 14, // px
        face: 'arial',
        vadjust: 0,
        mod: 'bold italic'
      },
      mono: {
        color: '#343434',
        size: 15, // px
        face: 'courier new',
        vadjust: 2,
        mod: ''
      }
    },
    group: undefined,
    heightConstraint: false,
    hidden: false,
    icon: {
      face: 'FontAwesome',
      code: undefined,
      size: 50,  //50,
      color:'#2B7CE9'
    },
    image: undefined,
    imagePadding: {
      left: 0,
      top: 0,
      bottom: 0,
      right: 0
    },
    label: undefined,
    labelHighlightBold: true,
    level: undefined,
    mass: 1,
    physics: true,
    scaling: {
      min: 10,
      max: 30,
      label: {
        enabled: false,
        min: 14,
        max: 30,
        maxVisible: 30,
        drawThreshold: 5
      },
      customScalingFunction: function (min,max,total,value) {
        if (max === min) {
          return 0.5;
        }
        else {
          let scale = 1 / (max - min);
          return Math.max(0,(value - min)*scale);
        }
      }
    },
    shadow:{
      enabled: false,
      color: 'rgba(0,0,0,0.5)',
      size:10,
      x:5,
      y:5
    },
    shape: 'ellipse',
    shapeProperties: {
      borderDashes: false, // only for borders
      borderRadius: 6,     // only for box shape
      interpolation: false,  // only for image and circularImage shapes
      useImageSize: false,  // only for image and circularImage shapes
      useBorderWithImage: false  // only for image shape
    }
    size: 25,
    title: undefined,
    value: undefined,
    widthConstraint: false,
    x: undefined,
    y: undefined
  }
}

network.setOptions(options)
#>
function New-DiagramOptionsInteraction {
    <#
    .SYNOPSIS
    Short description

    .DESCRIPTION
    Long description

    .PARAMETER DragNodes
    Parameter description

    .PARAMETER DragView
    Parameter description

    .PARAMETER HideEdgesOnDrag
    Parameter description

    .PARAMETER HideEdgesOnZoom
    Parameter description

    .PARAMETER HideNodesOnDrag
    Parameter description

    .PARAMETER Hover
    Parameter description

    .PARAMETER HoverConnectedEdges
    Parameter description

    .PARAMETER KeyboardEnabled
    Parameter description

    .PARAMETER KeyboardSpeedX
    Parameter description

    .PARAMETER KeyboardSpeedY
    Parameter description

    .PARAMETER KeyboardSpeedZoom
    Parameter description

    .PARAMETER KeyboardBindToWindow
    Parameter description

    .PARAMETER Multiselect
    Parameter description

    .PARAMETER NavigationButtons
    Parameter description

    .PARAMETER Selectable
    Parameter description

    .PARAMETER SelectConnectedEdges
    Parameter description

    .PARAMETER TooltipDelay
    Parameter description

    .PARAMETER ZoomView
    Parameter description

    .EXAMPLE
    An example

    .NOTES
    Based on options https://visjs.github.io/vis-network/docs/network/interaction.html#

    #>
    [alias('DiagramOptionsInteraction')]
    [CmdletBinding()]
    param(
        [nullable[bool]] $DragNodes,
        [nullable[bool]] $DragView,
        [nullable[bool]] $HideEdgesOnDrag,
        [nullable[bool]] $HideEdgesOnZoom,
        [nullable[bool]] $HideNodesOnDrag,
        [nullable[bool]] $Hover,
        [nullable[bool]] $HoverConnectedEdges,
        [nullable[bool]] $KeyboardEnabled,
        [nullable[int]] $KeyboardSpeedX,
        [nullable[int]] $KeyboardSpeedY,
        [nullable[decimal]] $KeyboardSpeedZoom,
        [nullable[bool]] $KeyboardBindToWindow,
        [nullable[bool]] $Multiselect,
        [nullable[bool]] $NavigationButtons,
        [nullable[bool]] $Selectable,
        [nullable[bool]] $SelectConnectedEdges,
        [nullable[int]] $TooltipDelay,
        [nullable[bool]] $ZoomView
    )

    $Object = [PSCustomObject] @{
        Type     = 'DiagramOptionsInteraction'
        Settings = @{
            interaction = [ordered] @{
                dragNodes            = $DragNodes
                dragView             = $DragView
                hideEdgesOnDrag      = $HideEdgesOnDrag
                hideEdgesOnZoom      = $HideEdgesOnZoom
                hideNodesOnDrag      = $HideNodesOnDrag
                hover                = $Hover
                hoverConnectedEdges  = $HoverConnectedEdges
                keyboard             = @{
                    enabled      = $KeyboardEnabled
                    speed        = @{
                        x    = $KeyboardSpeedX
                        y    = $KeyboardSpeedY
                        zoom = $KeyboardSpeedZoom
                    }
                    bindToWindow = $KeyboardBindToWindow
                }
                multiselect          = $Multiselect
                navigationButtons    = $NavigationButtons
                selectable           = $Selectable
                selectConnectedEdges = $SelectConnectedEdges
                tooltipDelay         = $TooltipDelay
                zoomView             = $ZoomView
            }
        }
    }
    Remove-EmptyValues -Hashtable $Object.Settings -Recursive -Rerun 2
    $Object
}
<#
    var options = {
        nodes: {
          borderWidth:2,
          borderWidthSelected: 8,
          size:24,
          color: {
            border: 'white',
            background: 'black',
            highlight: {
              border: 'black',
              background: 'white'
            },
            hover: {
              border: 'orange',
              background: 'grey'
            }
          },
          font:{color:'#eeeeee'},
          shapeProperties: {
            useBorderWithImage:true
          }
        },
        edges: {
          color: 'lightgray'
        }
      };
    #>


# https://visjs.github.io/vis-network/docs/network/edges.html#
function New-DiagramOptionsLayout {
    [alias('DiagramOptionsLayout')]
    [CmdletBinding()]
    param(
        [nullable[int]] $RandomSeed,
        [nullable[bool]] $ImprovedLayout,
        [nullable[int]] $ClusterThreshold ,
        [nullable[bool]] $HierarchicalEnabled,
        [nullable[int]] $HierarchicalLevelSeparation,
        [nullable[int]] $HierarchicalNodeSpacing,
        [nullable[int]] $HierarchicalTreeSpacing,
        [nullable[bool]] $HierarchicalBlockShifting,
        [nullable[bool]] $HierarchicalEdgeMinimization,
        [nullable[bool]] $HierarchicalParentCentralization,
        [ValidateSet('FromUpToDown', 'FromDownToUp', 'FromLeftToRight', 'FromRigthToLeft')][string] $HierarchicalDirection,
        [ValidateSet('hubsize', 'directed')][string] $HierarchicalSortMethod
    )
    $Direction = @{
        FromUpToDown    = 'UD'
        FromDownToUp    = 'DU'
        FromLeftToRight = 'LR'
        FromRigthToLeft = 'RL'
    }

    $Object = [PSCustomObject] @{
        Type     = 'DiagramOptionsLayout'
        Settings = @{
            layout = [ordered] @{
                randomSeed       = $RandomSeed
                improvedLayout   = $ImprovedLayout
                clusterThreshold = $ClusterThreshold
                hierarchical     = @{
                    enabled              = $HierarchicalEnabled
                    levelSeparation      = $HierarchicalLevelSeparation
                    nodeSpacing          = $HierarchicalNodeSpacing
                    treeSpacing          = $HierarchicalTreeSpacing
                    blockShifting        = $HierarchicalBlockShifting
                    edgeMinimization     = $HierarchicalEdgeMinimization
                    parentCentralization = $HierarchicalParentCentralization
                    direction            = $Direction[$HierarchicalDirection] # // UD, DU, LR, RL
                    sortMethod           = $HierarchicalSortMethod #// hubsize, directed
                }
            }
        }
    }
    Remove-EmptyValues -Hashtable $Object.Settings -Recursive -Rerun 2
    $Object
}

<#
// these are all options in full.
var options = {
  layout =  {
    randomSeed =  undefined,
    improvedLayout = true,
    clusterThreshold =  150,
    hierarchical =  {
      enabled = false,
      levelSeparation =  150,
      nodeSpacing =  100,
      treeSpacing =  200,
      blockShifting =  true,
      edgeMinimization =  true,
      parentCentralization =  true,
      direction =  'UD',        // UD, DU, LR, RL
      sortMethod =  'hubsize'   // hubsize, directed
    }
  }
}

network.setOptions(options);
#>
function New-DiagramOptionsLinks {
    [alias('DiagramOptionsEdges', 'New-DiagramOptionsEdges', 'DiagramOptionsLinks')]
    [CmdletBinding()]
    param(
        [nullable[bool]] $ArrowsToEnabled,
        [nullable[int]] $ArrowsToScaleFactor,
        [ValidateSet('arrow', 'bar', 'circle')][string] $ArrowsToType,
        [nullable[bool]] $ArrowsMiddleEnabled,
        [nullable[int]] $ArrowsMiddleScaleFactor,
        [ValidateSet('arrow', 'bar', 'circle')][string] $ArrowsMiddleType,
        [nullable[bool]] $ArrowsFromEnabled,
        [nullable[int]] $ArrowsFromScaleFactor,
        [ValidateSet('arrow', 'bar', 'circle')][string] $ArrowsFromType,
        [nullable[bool]] $ArrowStrikethrough,
        [nullable[bool]] $Chosen,
        [string] $Color,
        [string] $ColorHighlight,
        [string] $ColorHover,
        [ValidateSet('true', 'false', 'from', 'to', 'both')][string]$ColorInherit,
        [nullable[double]] $ColorOpacity, # range between 0 and 1
        [nullable[bool]]  $Dashes,
        [string] $Length,
        [string] $FontColor,
        [nullable[int]] $FontSize, #// px
        [string] $FontName,
        [string] $FontBackground,
        [nullable[int]] $FontStrokeWidth, #// px
        [string] $FontStrokeColor,
        [ValidateSet('center', 'left')][string] $FontAlign,
        [ValidateSet('false', 'true', 'markdown', 'html')][string]$FontMulti,
        [nullable[int]] $FontVAdjust,
        [nullable[int]] $WidthConstraint
    )
    $Object = [PSCustomObject] @{
        Type     = 'DiagramOptionsEdges'
        Settings = @{
            edges = [ordered] @{
                length             = $Length
                arrows             = [ordered]@{
                    to     = [ordered]@{
                        enabled     = $ArrowsToEnabled
                        scaleFactor = $ArrowsToScaleFactor
                        type        = $ArrowsToType
                    }
                    middle = [ordered]@{
                        enabled     = $ArrowsMiddleEnabled
                        scaleFactor = $ArrowsMiddleScaleFactor
                        type        = $ArrowsMiddleType
                    }
                    from   = [ordered]@{
                        enabled     = $ArrowsFromEnabled
                        scaleFactor = $ArrowsFromScaleFactor
                        type        = $ArrowsFromType
                    }
                }
                arrowStrikethrough = $ArrowStrikethrough
                chosen             = $Chosen
                color              = [ordered]@{
                    color     = ConvertFrom-Color -Color $Color
                    highlight = ConvertFrom-Color -Color $ColorHighlight
                    hover     = ConvertFrom-Color -Color $ColorHover
                    inherit   = $ColorInherit
                    opacity   = $ColorOpacity
                }
                font               = [ordered]@{
                    color       = ConvertFrom-Color -Color $FontColor
                    size        = $FontSize
                    face        = $FontName
                    background  = ConvertFrom-Color -Color $FontBackground
                    strokeWidth = $FontStrokeWidth
                    strokeColor = ConvertFrom-Color -Color $FontStrokeColor
                    align       = $FontAlign
                    multi       = $FontMulti
                    vadjust     = $FontVAdjust
                }
                dashes             = $Dashes
                widthConstraint    = $WidthConstraint
            }
        }
    }
    Remove-EmptyValues -Hashtable $Object.Settings -Recursive -Rerun 2
    $Object
}
Register-ArgumentCompleter -CommandName New-DiagramOptionsLinks -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramOptionsLinks -ParameterName ColorHighlight -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramOptionsLinks -ParameterName ColorHover -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramOptionsLinks -ParameterName FontColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramOptionsLinks -ParameterName FontBackground -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramOptionsLinks -ParameterName FontStrokeColor -ScriptBlock { $Script:RGBColors.Keys }
<#
// these are all options in full.
var options = {
  edges:{
    arrows: {
      to:     {enabled: false, scaleFactor:1, type:'arrow'},
      middle: {enabled: false, scaleFactor:1, type:'arrow'},
      from:   {enabled: false, scaleFactor:1, type:'arrow'}
    },
    arrowStrikethrough: true,
    chosen: true,
    color: {
      color:'#848484',
      highlight:'#848484',
      hover: '#848484',
      inherit: 'from',
      opacity:1.0
    },
    dashes: false,
    font: {
      color: '#343434',
      size: 14, // px
      face: 'arial',
      background: 'none',
      strokeWidth: 2, // px
      strokeColor: '#ffffff',
      align: 'horizontal',
      multi: false,
      vadjust: 0,
      bold: {
        color: '#343434',
        size: 14, // px
        face: 'arial',
        vadjust: 0,
        mod: 'bold'
      },
      ital: {
        color: '#343434',
        size: 14, // px
        face: 'arial',
        vadjust: 0,
        mod: 'italic',
      },
      boldital: {
        color: '#343434',
        size: 14, // px
        face: 'arial',
        vadjust: 0,
        mod: 'bold italic'
      },
      mono: {
        color: '#343434',
        size: 15, // px
        face: 'courier new',
        vadjust: 2,
        mod: ''
      }
    },
    hidden: false,
    hoverWidth: 1.5,
    label: undefined,
    labelHighlightBold: true,
    length: undefined,
    physics: true,
    scaling:{
      min: 1,
      max: 15,
      label: {
        enabled: true,
        min: 14,
        max: 30,
        maxVisible: 30,
        drawThreshold: 5
      },
      customScalingFunction: function (min,max,total,value) {
        if (max === min) {
          return 0.5;
        }
        else {
          var scale = 1 / (max - min);
          return Math.max(0,(value - min)*scale);
        }
      }
    },
    selectionWidth: 1,
    selfReferenceSize:20,
    shadow:{
      enabled: false,
      color: 'rgba(0,0,0,0.5)',
      size:10,
      x:5,
      y:5
    },
    smooth: {
      enabled: true,
      type: "dynamic",
      roundness: 0.5
    },
    title:undefined,
    value: undefined,
    width: 1,
    widthConstraint: false
  }
}

network.setOptions(options);
#>
function New-DiagramOptionsManipulation {
    <#
    .SYNOPSIS
    Short description

    .DESCRIPTION
    Long description

    .PARAMETER InitiallyActive
    Parameter description

    .PARAMETER AddNode
    Parameter description

    .PARAMETER AddEdge
    Parameter description

    .PARAMETER EditNode
    Parameter description

    .PARAMETER EditEdge
    Parameter description

    .PARAMETER DeleteNode
    Parameter description

    .PARAMETER DeleteEdge
    Parameter description

    .EXAMPLE
    An example

    .NOTES
    Based on https://visjs.github.io/vis-network/docs/network/manipulation.html#
    It's incomplete

    #>

    [alias('DiagramOptionsManipulation')]
    [CmdletBinding()]
    param(
        [nullable[bool]] $InitiallyActive,
        [nullable[bool]] $AddNode,
        [nullable[bool]] $AddEdge,
        [nullable[bool]] $EditNode,
        [nullable[bool]] $EditEdge,
        [nullable[bool]] $DeleteNode,
        [nullable[bool]] $DeleteEdge
    )

    $Object = [PSCustomObject] @{
        Type     = 'DiagramOptionsManipulation'
        Settings = @{
            manipulation = [ordered] @{
                enabled         = $true
                initiallyActive = $InitiallyActive
                addNode         = $AddNode
                addEdge         = $AddEdge
                editNode        = $EditNode
                editEdge        = $EditEdge
                deleteNode      = $DeleteNode
                deleteEdge      = $DeleteEdge
            }
        }
    }
    Remove-EmptyValues -Hashtable $Object.Settings -Recursive
    $Object
}
function New-DiagramOptionsNodes {
    [alias('DiagramOptionsNodes')]
    [CmdletBinding()]
    param(
        [nullable[int]] $BorderWidth,
        [nullable[int]] $BorderWidthSelected,
        [string] $BrokenImage,
        [nullable[bool]] $Chosen,
        [string] $ColorBorder,
        [string] $ColorBackground,
        [string] $ColorHighlightBorder,
        [string] $ColorHighlightBackground,
        [string] $ColorHoverBorder,
        [string] $ColorHoverBackground,
        [nullable[bool]] $FixedX,
        [nullable[bool]] $FixedY,
        [string] $FontColor,
        [nullable[int]] $FontSize, #// px
        [string] $FontName,
        [string] $FontBackground,
        [nullable[int]] $FontStrokeWidth, #// px
        [string] $FontStrokeColor,
        [ValidateSet('center', 'left')][string] $FontAlign,
        [ValidateSet('false', 'true', 'markdown', 'html')][string]$FontMulti,
        [nullable[int]] $FontVAdjust,
        [nullable[int]] $Size,
        [parameter(ParameterSetName = "Shape")][string][ValidateSet(
            'circle', 'dot', 'diamond', 'ellipse', 'database', 'box', 'square', 'triangle', 'triangleDown', 'text', 'star', 'hexagon')] $Shape,
        [nullable[int]] $HeightConstraintMinimum,
        [ValidateSet('top', 'middle', 'bottom')][string] $HeightConstraintVAlign,
        [nullable[int]] $WidthConstraintMinimum,
        [nullable[int]] $WidthConstraintMaximum,
        [nullable[int]] $Margin,
        [nullable[int]] $MarginTop,
        [nullable[int]] $MarginRight,
        [nullable[int]] $MarginBottom,
        [nullable[int]] $MarginLeft
    )
    $Object = [PSCustomObject] @{
        Type     = 'DiagramOptionsNodes'
        Settings = @{
            nodes = [ordered] @{
                borderWidth         = $BorderWidth
                borderWidthSelected = $BorderWidthSelected
                brokenImage         = $BrokenImage
                chosen              = $Chosen
                color               = [ordered]@{
                    border     = ConvertFrom-Color -Color $ColorBorder
                    background = ConvertFrom-Color -Color $ColorBackground
                    highlight  = [ordered]@{
                        border     = ConvertFrom-Color -Color $ColorHighlightBorder
                        background = ConvertFrom-Color -Color $ColorHighlightBackground
                    }
                    hover      = [ordered]@{
                        border     = ConvertFrom-Color -Color $ColorHoverBorder
                        background = ConvertFrom-Color -Color $ColorHoverBackground
                    }
                }
                fixed               = [ordered]@{
                    x = $FixedX
                    y = $FixedY
                }
                font                = [ordered]@{
                    color       = ConvertFrom-Color -Color $FontColor
                    size        = $FontSize #// px
                    face        = $FontName
                    background  = ConvertFrom-Color -Color $FontBackground
                    strokeWidth = $FontStrokeWidth #// px
                    strokeColor = ConvertFrom-Color -Color $FontStrokeColor
                    align       = $FontAlign
                    multi       = $FontMulti
                    vadjust     = $FontVAdjust
                }
                heightConstraint    = @{
                    minimum = $HeightConstraintMinimum
                    valign  = $HeightConstraintVAlign
                }
                size                = $Size
                shape               = $Shape
                widthConstraint     = @{
                    minimum = $WidthConstraintMinimum
                    maximum = $WidthConstraintMaximum
                }
            }
        }
    }

    if ($Margin) {
        $Object.Settings.nodes.margin = $Margin
    } else {
        $Object.Settings.nodes.margin = @{
            top    = $MarginTop
            right  = $MarginRight
            bottom = $MarginBottom
            left   = $MarginLeft
        }
    }


    Remove-EmptyValues -Hashtable $Object.Settings -Recursive -Rerun 2
    $Object
}

Register-ArgumentCompleter -CommandName New-DiagramOptionsNodes -ParameterName ColorBorder -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramOptionsNodes -ParameterName ColorBackground -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramOptionsNodes -ParameterName ColorHighlightBorder -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramOptionsNodes -ParameterName ColorHighlightBackground -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramOptionsNodes -ParameterName ColorHoverBorder -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramOptionsNodes -ParameterName ColorHoverBackground -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramOptionsNodes -ParameterName FontColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramOptionsNodes -ParameterName FontBackground -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-DiagramOptionsNodes -ParameterName FontStrokeColor -ScriptBlock { $Script:RGBColors.Keys }

<#
// these are all options in full.
var options = {
  nodes:{
    borderWidth: 1,
    borderWidthSelected: 2,
    brokenImage:undefined,
    chosen: true,
    color: {
      border: '#2B7CE9',
      background: '#97C2FC',
      highlight: {
        border: '#2B7CE9',
        background: '#D2E5FF'
      },
      hover: {
        border: '#2B7CE9',
        background: '#D2E5FF'
      }
    },
    fixed: {
      x:false,
      y:false
    },
    font: {
      color: '#343434',
      size: 14, // px
      face: 'arial',
      background: 'none',
      strokeWidth: 0, // px
      strokeColor: '#ffffff',
      align: 'center',
      multi: false,
      vadjust: 0,
      bold: {
        color: '#343434',
        size: 14, // px
        face: 'arial',
        vadjust: 0,
        mod: 'bold'
      },
      ital: {
        color: '#343434',
        size: 14, // px
        face: 'arial',
        vadjust: 0,
        mod: 'italic',
      },
      boldital: {
        color: '#343434',
        size: 14, // px
        face: 'arial',
        vadjust: 0,
        mod: 'bold italic'
      },
      mono: {
        color: '#343434',
        size: 15, // px
        face: 'courier new',
        vadjust: 2,
        mod: ''
      }
    },
    group: undefined,
    heightConstraint: false,
    hidden: false,
    icon: {
      face: 'FontAwesome',
      code: undefined,
      size: 50,  //50,
      color:'#2B7CE9'
    },
    image: undefined,
    imagePadding: {
      left: 0,
      top: 0,
      bottom: 0,
      right: 0
    },
    label: undefined,
    labelHighlightBold: true,
    level: undefined,
    mass: 1,
    physics: true,
    scaling: {
      min: 10,
      max: 30,
      label: {
        enabled: false,
        min: 14,
        max: 30,
        maxVisible: 30,
        drawThreshold: 5
      },
      customScalingFunction: function (min,max,total,value) {
        if (max === min) {
          return 0.5;
        }
        else {
          let scale = 1 / (max - min);
          return Math.max(0,(value - min)*scale);
        }
      }
    },
    shadow:{
      enabled: false,
      color: 'rgba(0,0,0,0.5)',
      size:10,
      x:5,
      y:5
    },
    shape: 'ellipse',
    shapeProperties: {
      borderDashes: false, // only for borders
      borderRadius: 6,     // only for box shape
      interpolation: false,  // only for image and circularImage shapes
      useImageSize: false,  // only for image and circularImage shapes
      useBorderWithImage: false  // only for image shape
    }
    size: 25,
    title: undefined,
    value: undefined,
    widthConstraint: false,
    x: undefined,
    y: undefined
  }
}

network.setOptions(options)
#>
function New-DiagramOptionsPhysics {
    [alias('DiagramOptionsPhysics')]
    [CmdletBinding()]
    param(
        [nullable[bool]] $Enabled,
        [nullable[bool]] $StabilizationEnabled,
        [nullable[int]] $Stabilizationiterations,
        [nullable[int]] $StabilizationupdateInterval,
        [nullable[bool]] $StabilizationonlyDynamicEdges,
        [nullable[bool]] $Stabilizationfit,
        [nullable[int]] $MaxVelocity,
        [nullable[int]] $MinVelocity,
        [nullable[int]] $Timestep,
        [nullable[bool]] $AdaptiveTimestep
    )
    $Object = [PSCustomObject] @{
        Type     = 'DiagramOptionsPhysics'
        Settings = @{
            physics = [ordered] @{
                enabled          = $Enabled
                stabilization    = @{
                    enabled          = $StabilizationEnabled
                    iterations       = $Stabilizationiterations
                    updateInterval   = $StabilizationupdateInterval
                    onlyDynamicEdges = $StabilizationonlyDynamicEdges
                    fit              = $Stabilizationfit
                }
                maxVelocity      = $MaxVelocity
                minVelocity      = $MinVelocity
                timestep         = $Timestep
                adaptiveTimestep = $AdaptiveTimestep
            }
        }
    }
    Remove-EmptyValues -Hashtable $Object.Settings -Recursive -Rerun 2
    $Object
}

<#
// these are all options in full.
var options = {
  physics:{
    enabled: true,
    barnesHut: {
      gravitationalConstant: -2000,
      centralGravity: 0.3,
      springLength: 95,
      springConstant: 0.04,
      damping: 0.09,
      avoidOverlap: 0
    },
    forceAtlas2Based: {
      gravitationalConstant: -50,
      centralGravity: 0.01,
      springConstant: 0.08,
      springLength: 100,
      damping: 0.4,
      avoidOverlap: 0
    },
    repulsion: {
      centralGravity: 0.2,
      springLength: 200,
      springConstant: 0.05,
      nodeDistance: 100,
      damping: 0.09
    },
    hierarchicalRepulsion: {
      centralGravity: 0.0,
      springLength: 100,
      springConstant: 0.01,
      nodeDistance: 120,
      damping: 0.09
    },
    maxVelocity: 50,
    minVelocity: 0.1,
    solver: 'barnesHut',
    stabilization: {
      enabled: true,
      iterations: 1000,
      updateInterval: 100,
      onlyDynamicEdges: false,
      fit: true
    },
    timestep: 0.5,
    adaptiveTimestep: true
  }
}

network.setOptions(options);
#>
function New-GageSector {
    [CmdletBinding()]
    param(
        [string] $Color,
        [int] $Min,
        [int] $Max
    )

    [ordered] @{
        color = ConvertFrom-Color -Color $Color
        lo    = $Min
        hi    = $Max
    }
}
Register-ArgumentCompleter -CommandName New-GageSection -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
function New-HierarchicalTreeNode {
    [alias('New-HierarchicalTreeNode', 'HierarchicalTreeNode')]
    [CmdletBinding()]
    param(
        [string] $ID,
        [alias('Name')][string] $Label,
        [string] $Type = "Organism",
        [string] $Description,
        [string] $To
    )

    if (-not $ID) {
        $ID = $Label
    }

    $Object = [PSCustomObject] @{
        Type     = 'TreeNode'
        Settings = [ordered] @{
            "id"          = $ID
            "parentId"    = $To
            "name"        = $Label
            #"type"        = $Type
            "description" = $Description
        }
    }
    Remove-EmptyValues -Hashtable $Object.Settings -Recursive
    $Object
}


Function New-HTML {
    [alias('Dashboard')]
    [CmdletBinding()]
    param(
        [alias('Content')][Parameter(Position = 0)][ValidateNotNull()][ScriptBlock] $HtmlData = $(Throw "Have you put the open curly brace on the next line?"),
        [switch] $UseCssLinks,
        [switch] $UseJavaScriptLinks,
        [alias('Name', 'Title')][String] $TitleText,
        [string] $Author,
        [string] $DateFormat = 'yyyy-MM-dd HH:mm:ss',
        [int] $AutoRefresh,
        # save HTML options
        [Parameter(Mandatory = $false)][string]$FilePath,
        [alias('Show', 'Open')][Parameter(Mandatory = $false)][switch]$ShowHTML,
        [ValidateSet('Unknown', 'String', 'Unicode', 'Byte', 'BigEndianUnicode', 'UTF8', 'UTF7', 'UTF32', 'Ascii', 'Default', 'Oem', 'BigEndianUTF32')] $Encoding = 'UTF8'
    )
    [string] $CurrentDate = (Get-Date).ToString($DateFormat)
    $Script:HTMLSchema = @{
        TabsHeaders       = [System.Collections.Generic.List[System.Collections.IDictionary]]::new() # tracks / stores headers
        TabsHeadersNested = [System.Collections.Generic.List[System.Collections.IDictionary]]::new() # tracks / stores headers
        Features          = [ordered] @{ } # tracks features for CSS/JS implementation
        Charts            = [System.Collections.Generic.List[string]]::new()
        Diagrams          = [System.Collections.Generic.List[string]]::new()

        Logos             = ""
        # Tab settings
        TabOptions        = @{
            SlimTabs = $false
        }

        CustomCSS         = [System.Collections.Generic.List[Array]]::new()
    }

    [Array] $TempOutputHTML = Invoke-Command -ScriptBlock $HtmlData

    $HeaderHTML = @()
    #$MainHTML = @()
    $FooterHTML = @()


    $MainHTML = foreach ($_ in $TempOutputHTML) {
        if ($_ -is [PSCustomObject]) {
            if ($_.Type -eq 'Footer') {
                $FooterHTML = $_.Output
            } elseif ($_.Type -eq 'Header') {
                $HeaderHTML = $_.Output
            } else {
                if ($_.Output) {
                    # this gets rid of any non-strings
                    # it's added here to track nested tabs
                    if ($_.Output -isnot [System.Collections.IDictionary]) {
                        $_.Output
                    }
                }
            }
        } else {
            # this gets rid of any non-strings
            # it's added here to track nested tabs
            if ($_ -isnot [System.Collections.IDictionary]) {
                $_
            }
        }
    }
    <#
    if ($MainHTML.Count -eq 0) {
        # this gets rid of any non-strings
        # it's added here to track nested tabs
        $MainHTML = foreach ($_ in $MainHTML) {
            if ($_ -isnot [System.Collections.IDictionary]) {
                $_
            }
        }
    } else {
        $MainHTML = foreach ($_ in $MainHTML) {
            if ($_ -isnot [System.Collections.IDictionary]) {
                $_
            }
        }
    }
    #>

    $Features = Get-FeaturesInUse -PriorityFeatures 'JQuery', 'DataTables', 'Tabs'


    # This removes Nested Tabs from primary Tabs
    foreach ($_ in $Script:HTMLSchema.TabsHeadersNested) {
        $null = $Script:HTMLSchema.TabsHeaders.Remove($_)
    }


    $HTML = @(
        '<!DOCTYPE html>'
        #"<!-- saved from url=(0016)http://localhost -->" + "`r`n"
        '<!-- saved from url=(0014)about:internet -->'
        New-HTMLTag -Tag 'html' {
            '<!-- HEAD -->'
            New-HTMLTag -Tag 'head' {
                New-HTMLTag -Tag 'meta' -Attributes @{ charset = "utf-8" } -SelfClosing
                #New-HTMLTag -Tag 'meta' -Attributes @{ 'http-equiv' = 'X-UA-Compatible'; content = 'IE=8' } -SelfClosing
                New-HTMLTag -Tag 'meta' -Attributes @{ name = 'viewport'; content = 'width=device-width, initial-scale=1' } -SelfClosing
                New-HTMLTag -Tag 'meta' -Attributes @{ name = 'author'; content = $Author } -SelfClosing
                New-HTMLTag -Tag 'meta' -Attributes @{ name = 'revised'; content = $CurrentDate } -SelfClosing
                New-HTMLTag -Tag 'title' { $TitleText }
                if ($Autorefresh -gt 0) {
                    New-HTMLTag -Tag 'meta' -Attributes @{ 'http-equiv' = 'refresh'; content = $Autorefresh } -SelfClosing
                }
                Get-Resources -UseCssLinks:$true -UseJavaScriptLinks:$true -Location 'HeaderAlways' -Features Default, DefaultHeadings, Fonts, FontsAwesome
                Get-Resources -UseCssLinks:$false -UseJavaScriptLinks:$false -Location 'HeaderAlways' -Features Default, DefaultHeadings
                if ($null -ne $Features) {
                    Get-Resources -UseCssLinks:$true -UseJavaScriptLinks:$true -Location 'HeaderAlways' -Features $Features -NoScript
                    Get-Resources -UseCssLinks:$false -UseJavaScriptLinks:$false -Location 'HeaderAlways' -Features $Features -NoScript
                    Get-Resources -UseCssLinks:$UseCssLinks -UseJavaScriptLinks:$UseJavaScriptLinks -Location 'Header' -Features $Features
                }
            }

            New-HTMLCustomCSS -Css $Script:HTMLSchema.CustomCSS
            '<!-- END HEAD -->'
            '<!-- BODY -->'
            New-HTMLTag -Tag 'body' {
                '<!-- HEADER -->'
                New-HTMLTag -Tag 'header' {
                    if ($HeaderHTML) {
                        $HeaderHTML
                    }
                }
                '<!-- END HEADER -->'
                # Add logo if there is one
                $Script:HTMLSchema.Logos
                # Add tabs header if there is one
                if ($Script:HTMLSchema.TabsHeaders) {
                    New-HTMLTabHead
                    New-HTMLTag -Tag 'div' -Attributes @{ 'data-panes' = 'true' } {
                        # Add remaining data
                        #$OutputHTML
                        $MainHTML
                    }
                } else {
                    # Add remaining data
                    $MainHTML
                    #$OutputHTML
                }
                # Add charts scripts if those are there
                foreach ($Chart in $Script:HTMLSchema.Charts) {
                    $Chart
                }
                foreach ($Diagram in $Script:HTMLSchema.Diagrams) {
                    $Diagram
                }

                New-HTMLTag -Tag 'footer' {
                    '<!-- FOOTER -->'
                    if ($FooterHTML) {
                        $FooterHTML
                    }
                    #New-HTMLTag -Tag 'footer' {
                    if ($null -ne $Features) {
                        # FooterAlways means we're not able to provide consistent output with and without links and we prefer those to be included
                        # either as links or from file per required features
                        Get-Resources -UseCssLinks:$true -UseJavaScriptLinks:$true -Location 'FooterAlways' -Features $Features
                        Get-Resources -UseCssLinks:$false -UseJavaScriptLinks:$false -Location 'FooterAlways' -Features $Features
                        # standard footer features
                        Get-Resources -UseCssLinks:$UseCssLinks -UseJavaScriptLinks:$UseJavaScriptLinks -Location 'Footer' -Features $Features
                    }
                    '<!-- END FOOTER -->'
                }
                '<!-- END BODY -->'
            }

        }
    )
    if ($FilePath -ne '') {
        Save-HTML -HTML $HTML -FilePath $FilePath -ShowHTML:$ShowHTML -Encoding $Encoding
    } else {
        $HTML
    }
}
function New-HTMLCalendar {
    [alias('Calendar')]
    [CmdletBinding()]
    param(
        [ScriptBlock] $CalendarSettings,
        [ValidateSet('interaction', 'dayGrid', 'timeGrid', 'list', 'rrule')][string[]] $Plugins = @('interaction', 'dayGrid', 'timeGrid', 'list', 'rrule'),
        [ValidateSet(
            'prev', 'next', 'today', 'prevYear', 'nextYear', 'dayGridDay', 'dayGridWeek', 'dayGridMonth',
            'timeGridWeek', 'timeGridDay', 'listDay', 'listWeek', 'listMonth', 'title'
        )][string[]] $HeaderLeft = @('prev', 'next', 'today'),
        [ValidateSet(
            'prev', 'next', 'today', 'prevYear', 'nextYear', 'dayGridDay', 'dayGridWeek', 'dayGridMonth',
            'timeGridWeek', 'timeGridDay', 'listDay', 'listWeek', 'listMonth', 'title'
        )][string[]]$HeaderCenter = 'title',
        [ValidateSet(
            'prev', 'next', 'today', 'prevYear', 'nextYear', 'dayGridDay', 'dayGridWeek', 'dayGridMonth',
            'timeGridWeek', 'timeGridDay', 'listDay', 'listWeek', 'listMonth', 'title'
        )][string[]] $HeaderRight = @('dayGridMonth', 'timeGridWeek', 'timeGridDay', 'listMonth'),
        [DateTime] $DefaultDate = (Get-Date),
        [bool] $NavigationLinks = $true,
        [bool] $NowIndicator = $true,
        [bool] $EventLimit = $true,
        [bool] $WeekNumbers = $true,
        [bool] $WeekNumbersWithinDays = $true,
        [bool] $Selectable = $true,
        [bool] $SelectMirror = $true,
        [switch] $BusinessHours,
        [switch] $Editable
    )
    if (-not $Script:HTMLSchema.Features) {
        Write-Warning 'New-HTMLCalendar - Creation of HTML aborted. Most likely New-HTML is missing.'
        Exit
    }
    $Script:HTMLSchema.Features.FullCalendar = $true
    $Script:HTMLSchema.Features.FullCalendarCore = $true
    $Script:HTMLSchema.Features.FullCalendarDayGrid = $true
    $Script:HTMLSchema.Features.FullCalendarInteraction = $true
    $Script:HTMLSchema.Features.FullCalendarList = $true
    $Script:HTMLSchema.Features.FullCalendarRRule = $true
    $Script:HTMLSchema.Features.FullCalendarTimeGrid = $true
    $Script:HTMLSchema.Features.FullCalendarTimeLine = $true
    $Script:HTMLSchema.Features.Popper = $true

    $CalendarEvents = [System.Collections.Generic.List[System.Collections.IDictionary]]::new()

    [Array] $Settings = & $CalendarSettings
    foreach ($Object in $Settings) {
        if ($Object.Type -eq 'CalendarEvent') {
            $CalendarEvents.Add($Object.Settings)
        }
    }

    # Define HTML/Script
    [string] $ID = "Calendar-" + (Get-RandomStringName -Size 8)

    $Calendar = [ordered] @{
        plugins               = $Plugins
        header                = @{
            left   = $HeaderLeft -join ','
            center = $HeaderCenter -join ','
            right  = $HeaderRight -join ','
        }
        defaultDate           = '{0:yyyy-MM-dd}' -f ($DefaultDate)
        nowIndicator          = $NowIndicator
        #now: '2018-02-13T09:25:00' // just for demo
        navLinks              = $NavigationLinks #// can click day/week names to navigate views
        businessHours         = $BusinessHours.IsPresent #// display business hours
        editable              = $Editable.IsPresent
        events                = $CalendarEvents
        eventLimit            = $EventLimit
        weekNumbers           = $WeekNumbers
        weekNumbersWithinDays = $WeekNumbersWithinDays
        weekNumberCalculation = 'ISO'
        selectable            = $Selectable
        selectMirror          = $SelectMirror
        buttonIcons           = $false # // show the prev/next text
        #// customize the button names,
        #// otherwise they'd all just say "list"
        views                 = @{
            listDay   = @{ buttonText = 'list day' }
            listWeek  = @{ buttonText = 'list week' }
            listMonth = @{ buttonText = 'list month' }
        }
        eventRender           = 'ReplaceMe'
    }
    Remove-EmptyValues -Hashtable $Calendar -Recursive
    $CalendarJSON = $Calendar | ConvertTo-Json -Depth 7

    # Adding function for ToolTips / need cleaner way
    $EventRender = @"
    eventRender: function (info) {
        var tooltip = new Tooltip(info.el, {
            title: info.event.extendedProps.description,
            placement: 'top',
            trigger: 'hover',
            container: 'body'
        });
    }
"@
    if ($PSEdition -eq 'Desktop') {
        $TextToFind = '"eventRender":  "ReplaceMe"'
    } else {
        $TextToFind = '"eventRender": "ReplaceMe"'
    }
    $CalendarJSON = $CalendarJSON.Replace($TextToFind, $EventRender)

    $Div = New-HTMLTag -Tag 'div' -Attributes @{ id = $ID; class = 'calendarFullCalendar'; style = $Style }
    $Script = New-HTMLTag -Tag 'script' -Value {
        "document.addEventListener('DOMContentLoaded', function () {"
        "var calendarEl = document.getElementById('$ID');"
        'var calendar = new FullCalendar.Calendar(calendarEl,'
        $CalendarJSON
        ');'
        'calendar.render();'
        '}); '
    } -NewLine

    # return HTML
    $Script
    $Div
}
function New-HTMLChart {
    [alias('Chart')]
    [CmdletBinding()]
    param(
        [ScriptBlock] $ChartSettings,
        [string] $Title,
        [ValidateSet('center', 'left', 'right', 'default')][string] $TitleAlignment = 'default',
        [nullable[int]] $Height = 350,
        [nullable[int]] $Width,
        [alias('GradientColors')][switch] $Gradient,
        [alias('PatternedColors')][switch] $Patterned
    )

    # Datasets Bar/Line
    $DataSet = [System.Collections.Generic.List[object]]::new()
    $DataName = [System.Collections.Generic.List[object]]::new()


    # Legend Variables
    $Colors = [System.Collections.Generic.List[string]]::new()

    # Line Variables
    # $LineColors = [System.Collections.Generic.List[string]]::new()
    $LineCurves = [System.Collections.Generic.List[string]]::new()
    $LineWidths = [System.Collections.Generic.List[int]]::new()
    $LineDashes = [System.Collections.Generic.List[int]]::new()
    $LineCaps = [System.Collections.Generic.List[string]]::new()

    #$RadialColors = [System.Collections.Generic.List[string]]::new()
    #$SparkColors = [System.Collections.Generic.List[string]]::new()

    # Bar default definitions
    [bool] $BarHorizontal = $true
    [bool] $BarDataLabelsEnabled = $true
    [int] $BarDataLabelsOffsetX = -6
    [string] $BarDataLabelsFontSize = '12px'
    [bool] $BarDistributed = $false

    [string] $LegendPosition = 'default'
    #
    [string] $Type = ''

    [Array] $Settings = & $ChartSettings
    foreach ($Setting in $Settings) {
        if ($Setting.ObjectType -eq 'Bar') {
            # For Bar Charts
            if (-not $Type) {
                # thiss makes sure type is not set if BarOptions is used which already set type to BarStacked or similar
                $Type = $Setting.ObjectType
            }
            $DataSet.Add($Setting.Value)
            $DataName.Add($Setting.Name)

        } elseif ($Setting.ObjectType -eq 'Pie' -or $Setting.ObjectType -eq 'Donut') {
            # For Pie Charts
            $Type = $Setting.ObjectType
            $DataSet.Add($Setting.Value)
            $DataName.Add($Setting.Name)

            if ($Setting.Color) {
                $Colors.Add($Setting.Color)
            }
        } elseif ($Setting.ObjectType -eq 'Spark') {
            # For Spark Charts
            $Type = $Setting.ObjectType
            $DataSet.Add($Setting.Value)
            $DataName.Add($Setting.Name)

            if ($Setting.Color) {
                $Colors.Add($Setting.Color)
            }
        } elseif ($Setting.ObjectType -eq 'Radial') {
            $Type = $Setting.ObjectType
            $DataSet.Add($Setting.Value)
            $DataName.Add($Setting.Name)

            if ($Setting.Color) {
                $Colors.Add($Setting.Color)
            }
        } elseif ($Setting.ObjectType -eq 'Legend') {
            # For Bar Charts
            $DataLegend = $Setting.Names
            $LegendPosition = $Setting.LegendPosition
            if ($null -ne $Setting.Color) {
                $Colors = $Setting.Color
            }
        } elseif ($Setting.ObjectType -eq 'BarOptions') {
            # For Bar Charts
            $Type = $Setting.Type
            $BarHorizontal = $Setting.Horizontal
            $BarDataLabelsEnabled = $Setting.DataLabelsEnabled
            $BarDataLabelsOffsetX = $Setting.DataLabelsOffsetX
            $BarDataLabelsFontSize = $Setting.DataLabelsFontSize
            $BarDataLabelsColor = $Setting.DataLabelsColor
            $BarDistributed = $Setting.Distributed

            # This is required to support legacy ChartBarOptions - Gradient -Patterned
            if ($null -ne $Setting.PatternedColors) {
                $Patterned = $Setting.PatternedColors
            }
            if ($null -ne $Setting.GradientColors) {
                $Gradient = $Setting.GradientColors
            }
        } elseif ($Setting.ObjectType -eq 'Toolbar') {
            # For All Charts
            $Toolbar = $Setting.Toolbar
        } elseif ($Setting.ObjectType -eq 'Theme') {
            # For All Charts
            $Theme = $Setting.Theme
        } elseif ($Setting.ObjectType -eq 'Line') {
            # For Line Charts
            $Type = $Setting.ObjectType
            $DataSet.Add($Setting.Value)
            $DataName.Add($Setting.Name)
            if ($Setting.LineColor) {
                $Colors.Add($Setting.LineColor)
            }
            if ($Setting.LineCurve) {
                $LineCurves.Add($Setting.LineCurve)
            }
            if ($Setting.LineWidth) {
                $LineWidths.Add($Setting.LineWidth)
            }
            if ($Setting.LineDash) {
                $LineDashes.Add($Setting.LineDash)
            }
            if ($Setting.LineCap) {
                $LineCaps.Add($Setting.LineCap)
            }
        } elseif ($Setting.ObjectType -eq 'ChartAxisX') {
            $ChartAxisX = $Setting.ChartAxisX
            #$DataCategory = $ChartAxisX.Names

        } elseif ($Setting.ObjectType -eq 'ChartGrid') {
            $GridOptions = $Setting.Grid
        } elseif ($Setting.ObjectType -eq 'ChartAxisY') {
            $ChartAxisY = $Setting.ChartAxisY
        }
    }

    if ($Type -in @('bar', 'barStacked', 'barStacked100Percent')) {
        if ($DataLegend.Count -lt $DataSet[0].Count) {
            Write-Warning -Message "Chart Legend count doesn't match values count. Skipping."
        }
        # Fixes dataset/dataname to format expected by New-HTMLChartBar
        $HashTable = [ordered] @{ }
        $ArrayCount = $DataSet[0].Count
        if ($ArrayCount -eq 1) {
            $HashTable.1 = $DataSet
        } else {
            for ($i = 0; $i -lt $ArrayCount; $i++) {
                $HashTable.$i = [System.Collections.Generic.List[object]]::new()
            }
            foreach ($Value in $DataSet) {
                for ($h = 0; $h -lt $Value.Count; $h++) {
                    $HashTable[$h].Add($Value[$h])
                }
            }
        }

        New-HTMLChartBar `
            -Data $($HashTable.Values) `
            -DataNames $DataName `
            -DataLegend $DataLegend `
            -LegendPosition $LegendPosition `
            -Type $Type `
            -Title $Title `
            -TitleAlignment $TitleAlignment `
            -Horizontal:$BarHorizontal `
            -DataLabelsEnabled $BarDataLabelsEnabled `
            -DataLabelsOffsetX $BarDataLabelsOffsetX `
            -DataLabelsFontSize $BarDataLabelsFontSize `
            -Distributed:$BarDistributed `
            -DataLabelsColor $BarDataLabelsColor `
            -Height $Height `
            -Width $Width `
            -Colors $Colors `
            -Theme $Theme -Toolbar $Toolbar -GridOptions $GridOptions -PatternedColors:$Patterned -GradientColors:$Gradient
    } elseif ($Type -eq 'Line') {
        if (-not $ChartAxisX) {
            Write-Warning -Message 'Chart Category (Chart Axis X) is missing.'
            Exit
        }
        New-HTMLChartLine -Data $DataSet `
            -DataNames $DataName `
            -DataLabelsEnabled $BarDataLabelsEnabled `
            -DataLabelsOffsetX $BarDataLabelsOffsetX `
            -DataLabelsFontSize $BarDataLabelsFontSize `
            -DataLabelsColor $BarDataLabelsColor `
            -LineColor $Colors `
            -LineCurve $LineCurves `
            -LineWidth $LineWidths `
            -LineDash $LineDashes `
            -LineCap $LineCaps `
            -ChartAxisX $ChartAxisX `
            -ChartAxisY $ChartAxisY `
            -Title $Title -TitleAlignment $TitleAlignment `
            -Height $Height -Width $Width `
            -Theme $Theme -Toolbar $Toolbar -GridOptions $GridOptions -PatternedColors:$Patterned -GradientColors:$Gradient

    } elseif ($Type -eq 'Pie' -or $Type -eq 'Donut') {
        New-HTMLChartPie `
            -Type $Type `
            -Data $DataSet `
            -DataNames $DataName `
            -Colors $Colors `
            -Title $Title -TitleAlignment $TitleAlignment `
            -Height $Height -Width $Width `
            -Theme $Theme -Toolbar $Toolbar -GridOptions $GridOptions -PatternedColors:$Patterned -GradientColors:$Gradient
    } elseif ($Type -eq 'Spark') {
        New-HTMLChartSpark `
            -Data $DataSet `
            -DataNames $DataName `
            -Colors $Colors `
            -Title $Title -TitleAlignment $TitleAlignment `
            -Height $Height -Width $Width `
            -Theme $Theme -Toolbar $Toolbar -GridOptions $GridOptions -PatternedColors:$Patterned -GradientColors:$Gradient
    } elseif ($Type -eq 'Radial') {
        New-HTMLChartRadial `
            -Data $DataSet `
            -DataNames $DataName `
            -Colors $Colors `
            -Title $Title -TitleAlignment $TitleAlignment `
            -Height $Height -Width $Width `
            -Theme $Theme -Toolbar $Toolbar -GridOptions $GridOptions -PatternedColors:$Patterned -GradientColors:$Gradient
    }
}
Function New-HTMLCodeBlock {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)][String] $Code,
        [Parameter(Mandatory = $false)]
        [ValidateSet(
            'assembly',
            'asm',
            'avrassembly',
            'avrasm',
            'c',
            'cpp',
            'c++',
            'csharp',
            'css',
            'cython',
            'cordpro',
            'diff',
            'docker',
            'dockerfile',
            'generic',
            'standard',
            'groovy',
            'go',
            'golang',
            'html',
            'ini',
            'conf',
            'java',
            'js',
            'javascript',
            'jquery',
            'mootools',
            'ext.js',
            'json',
            'kotlin',
            'less',
            'lua',
            'gfm',
            'md',
            'markdown',
            'octave',
            'matlab',
            'nsis',
            'php',
            'powershell',
            'prolog',
            'py',
            'python',
            'raw',
            'ruby',
            'rust',
            'scss',
            'sass',
            'shell',
            'bash',
            'sql',
            'squirrel',
            'swift',
            'typescript',
            'vhdl',
            'visualbasic',
            'vb',
            'xml',
            'yaml'
        )]

        [String] $Style = 'powershell',
        [Parameter(Mandatory = $false)]
        [ValidateSet(
            'enlighter',
            'beyond',
            'classic',
            'godzilla',
            'atomic',
            'droide',
            'minimal',
            'eclipse',
            'mowtwo',
            'rowhammer',
            'bootstrap4',
            'dracula',
            'monokai'
        )][String] $Theme,
        [Parameter(Mandatory = $false)][String] $Group,
        [Parameter(Mandatory = $false)][String] $Title,
        [Parameter(Mandatory = $false)][String[]] $Highlight,
        [Parameter(Mandatory = $false)][nullable[bool]] $ShowLineNumbers,
        [Parameter(Mandatory = $false)][String] $LineOffset
    )
    $Script:HTMLSchema.Features.CodeBlocks = $true
    <# Explanation to fields:
        data-enlighter-language (string) - The language of the codeblock - overrides the global default setting | Block+Inline Content option
        data-enlighter-theme (string) - The theme of the codeblock - overrides the global default setting | Block+Inline Content option
        data-enlighter-group (string) - The identifier of the codegroup where the codeblock belongs to | Block Content option
        data-enlighter-title (string) - The title/name of the tab | Block Content option
        data-enlighter-linenumbers (boolean) - Show/Hide the linenumbers of a codeblock (Values: "true", "false") | Block Content option
        data-enlighter-highlight (string) - A List of lines to point out, comma seperated (ranges are supported) e.g. "2,3,6-10" | Block Content option
        data-enlighter-lineoffset (number) - Start value of line-numbering e.g. "5" to start with line 5 - attribute start of the ol tag is set | Block Content option
    #>

    if ($null -eq $ShowLineNumbers -and $Highlight) {
        $ShowLineNumbers = $true
    }

    $Attributes = [ordered]@{
        'data-enlighter-language'    = "$Style".ToLower()
        'data-enlighter-theme'       = "$Theme".ToLower()
        'data-enlighter-group'       = "$Group".ToLower()
        'data-enlighter-title'       = "$Title"
        'data-enlighter-linenumbers' = "$ShowLineNumbers"
        'data-enlighter-highlight'   = "$Highlight"
        'data-enlighter-lineoffset'  = "$LineOffset".ToLower()
    }

    # Cleanup code (if there are spaces before code it fixes that)
    $ExtraCode = $Code.Split([System.Environment]::NewLine)
    [int] $Length = 5000
    $NewCode = foreach ($Line in $ExtraCode) {
        if ($Line.Trim() -ne '') {
            [int] $TempLength = $Line.Length - (($Line -replace '^(\s+)').Length)
            #$TempLength = ($line -replace '^(\s+).+$', '$1').Length
            if ($TempLength -le $Length) {
                $Length = $TempLength
            }
            $Line
        }
    }
    $FixedCode = foreach ($Line in $NewCode) {
        $Line.Substring($Length)
    }
    $FinalCode = $FixedCode -join [System.Environment]::NewLine
    # Prepare HTML
    New-HTMLTag -Tag 'pre' -Attributes $Attributes {
        $FinalCode
    }
}

function New-HTMLContainer {
    [alias('Container')]
    [CmdletBinding()]
    param(
        [alias('Content')][Parameter(Mandatory = $false, Position = 0)][ScriptBlock] $HTML,
        [string] $Width,
        [string] $Margin
    )

    if ($Width -or $Margin) {
        [string] $ClassName = "flexElement$(Get-RandomStringName -Size 8 -LettersOnly)"
        $Attributes = @{
            'flex-basis' = if ($Width) { $Width } else { '100%' }
            'margin'     = if ($Margin) { $Margin }
        }
        $Css = ConvertTo-CSS -ClassName $ClassName -Attributes $Attributes

        $Script:HTMLSchema.CustomCSS.Add($Css)
        [string] $Class = "$ClassName overflowHidden"
    } else {
        [string] $Class = 'flexElement overflowHidden'
    }
    New-HTMLTag -Tag 'div' -Attributes @{ class = $Class } {
        if ($HTML) {
            Invoke-Command -ScriptBlock $HTML
        }
    }
}
function New-HTMLDiagram {
    [alias('Diagram')]
    [CmdletBinding()]
    param(
        [ScriptBlock] $Diagram,
        [string] $Height,
        [string] $Width,
        [switch] $BundleImages,
        [uri] $BackGroundImage,
        [string] $BackgroundSize = '100% 100%'
    )
    if (-not $Script:HTMLSchema.Features) {
        Write-Warning 'New-HTMLDiagram - Creation of HTML aborted. Most likely New-HTML is missing.'
        Exit
    }

    $DataEdges = [System.Collections.Generic.List[System.Collections.IDictionary]]::new()
    $DataNodes = [ordered] @{ }
    $DataEvents = [System.Collections.Generic.List[System.Collections.IDictionary]]::new()

    [Array] $Settings = & $Diagram
    foreach ($Node in $Settings) {
        if ($Node.Type -eq 'DiagramNode') {
            $ID = $Node.Settings['id']
            $DataNodes[$ID] = $Node.Settings
            #$DataEdges.Add($Node.Edges)
            foreach ($From in $Node.Edges.From) {
                foreach ($To in $Node.Edges.To) {
                    $Edge = $Node.Edges.Clone()
                    $Edge['from'] = $From
                    $Edge['to'] = $To
                    $DataEdges.Add($Edge)
                }
            }
        } elseif ($Node.Type -eq 'DiagramOptionsInteraction') {
            $DiagramOptionsInteraction = $Node.Settings
        } elseif ($Node.Type -eq 'DiagramOptionsManipulation') {
            $DiagramOptionsManipulation = $Node.Settings
        } elseif ($Node.Type -eq 'DiagramOptionsPhysics') {
            $DiagramOptionsPhysics = $Node.Settings
        } elseif ($Node.Type -eq 'DiagramOptionsLayout') {
            $DiagramOptionsLayout = $Node.Settings
        } elseif ($Node.Type -eq 'DiagramOptionsNodes') {
            $DiagramOptionsNodes = $Node.Settings
        } elseif ($Node.Type -eq 'DiagramOptionsEdges') {
            $DiagramOptionsEdges = $Node.Settings
        } elseif ($Node.Type -eq 'DiagramLink') {
            if ($Node.Settings.From -and $Node.Settings.To) {
                foreach ($From in $Node.Settings.From) {
                    foreach ($To in $Node.Settings.To) {
                        $Edge = $Node.Edges.Clone()
                        $Edge['from'] = $From
                        $Edge['to'] = $To
                        $DataEdges.Add($Edge)
                    }
                }
            }
            $DataEdges.Add($Node.Edges)
        } elseif ($Node.Type -eq 'DiagramEvent') {
            $DataEvents.Add($Node.Settings)
        }
    }
    <#
    {id: 14, shape: 'circularImage', image: DIR + '14.png'},
    {id: 15, shape: 'circularImage', image: DIR + 'missing.png', brokenImage: DIR + 'missingBrokenImage.png', label:"when images\nfail\nto load"},
    {id: 16, shape: 'circularImage', image: DIR + 'anotherMissing.png', brokenImage: DIR + '9.png', label:"fallback image in action"}
    {id: 5, label:'colorObject', color: {background:'pink', border:'purple'}},
    {id: 6, label:'colorObject + highlight', color: {background:'#F03967', border:'#713E7F',highlight:{background:'red',border:'black'}}},
    {id: 7, label:'colorObject + highlight + hover', color: {background:'cyan', border:'blue',highlight:{background:'red',border:'blue'},hover:{background:'white',border:'red'}}}
    {id: 1,label: 'User 1',group: 'users'},
    {id: 2,label: 'User 2',group: 'users'},
    {id: 3,label: 'Usergroup 1',group: 'usergroups'}
    nodes.push({id: 1, label: 'Main', image: DIR + 'Network-Pipe-icon.png', shape: 'image'});
    nodes.push({id: 2, label: 'Office', image: DIR + 'Network-Pipe-icon.png', shape: 'image'});
    nodes.push({id: 3, label: 'Wireless', image: DIR + 'Network-Pipe-icon.png', shape: 'image'});
    {id: 3,  shape: 'image', image: DIR + '3.png', label: "imagePadding{2,10,8,20}+size", imagePadding: { left: 2, top: 10, right: 8, bottom: 20}, size: 40, color: { border: 'green', background: 'yellow', highlight: { border: 'yellow', background: 'green' }, hover: { border: 'orange', background: 'grey' } } },
    {id: 9,  shape: 'image', image: DIR + '9.png', label: "useImageSize + imagePadding:15", shapeProperties: { useImageSize: true }, imagePadding: 30, color: { border: 'blue', background: 'orange', highlight: { border: 'orange', background: 'blue' }, hover: { border: 'orange', background: 'grey' } } },
    var url = "data:image/svg+xml;charset=utf-8,"+ encodeURIComponent(svg);
    {id: 2, label: 'Using SVG', image: url, shape: 'image'}

    #>


    [Array] $Nodes = foreach ($_ in $DataNodes.Keys) {
        if ($DataNodes[$_]['image']) {
            if ($BundleImages) {
                $DataNodes[$_]['image'] = Convert-Image -Image $DataNodes[$_]['image']
            }
        }
        $NodeJson = $DataNodes[$_] | ConvertTo-Json -Depth 5 #| ForEach-Object { [System.Text.RegularExpressions.Regex]::Unescape($_) }

        # We need to fix wrong escaped chars, Unescape breaks other parts
        $Replace = @{
            '"\"Font Awesome 5 Solid\""'        = "'`"Font Awesome 5 Solid`"'"
            '"\"Font Awesome 5 Brands\""'       = "'`"Font Awesome 5 Brands`"'"
            '"\"Font Awesome 5 Regular\""'      = "'`"Font Awesome 5 Regular`"'"
            '"\"Font Awesome 5 Free\""'         = "'`"Font Awesome 5 Free`"'"
            '"\"Font Awesome 5 Free Regular\""' = "'`"Font Awesome 5 Free Regular`"'"
            '"\"Font Awesome 5 Free Solid\""'   = "'`"Font Awesome 5 Free Solid`"'"
            '"\"Font Awesome 5 Free Brands\""'  = "'`"Font Awesome 5 Free Brands`"'"
            '"\\u'                              = '"\u'
        }
        foreach ($R in $Replace.Keys) {
            $NodeJson = $NodeJson.Replace($R, $Replace[$R])
        }
        $NodeJson
    }
    [Array] $Edges = foreach ($_ in $DataEdges) {
        #if ($_.From -and $_.To) {
        #    foreach ($SingleTo in $_.To) {
        # [ordered] @{
        #    from = $_.From
        #     to   = $SingleTo
        # } | ConvertTo-Json -Depth 5
        $_ | ConvertTo-Json -Depth 5 #| ForEach-Object { [System.Text.RegularExpressions.Regex]::Unescape($_) }
        #}
        #}
    }

    $Options = @{ }
    if ($DiagramOptionsInteraction) {
        if ($DiagramOptionsInteraction['interaction']) {
            $Options['interaction'] = $DiagramOptionsInteraction['interaction']
        }
    }
    if ($DiagramOptionsManipulation) {
        if ($DiagramOptionsManipulation['manipulation']) {
            $Options['manipulation'] = $DiagramOptionsManipulation['manipulation']
        }
    }
    if ($DiagramOptionsPhysics) {
        if ($DiagramOptionsPhysics['physics']) {
            $Options['physics'] = $DiagramOptionsPhysics['physics']
        }
    }
    if ($DiagramOptionsLayout) {
        if ($DiagramOptionsLayout['layout']) {
            $Options['layout'] = $DiagramOptionsLayout['layout']
        }
    }
    if ($DiagramOptionsEdges) {
        if ($DiagramOptionsEdges['edges']) {
            $Options['edges'] = $DiagramOptionsEdges['edges']
        }
    }
    if ($DiagramOptionsNodes) {
        if ($DiagramOptionsNodes['nodes']) {
            $Options['nodes'] = $DiagramOptionsNodes['nodes']
        }
    }

    if ($BundleImages -and $BackGroundImage) {
        $Image = Convert-Image -Image $BackGroundImage
    } else {
        $Image = $BackGroundImage
    }

    New-InternalDiagram -Nodes $Nodes -Edges $Edges -Options $Options -Width $Width -Height $Height -BackgroundImage $Image -Events $DataEvents
}

function New-HTMLFooter {
    [alias('Footer')]
    [CmdletBinding()]
    param(
        [scriptblock] $HTMLContent
    )
    if ($HTMLContent) {
        [PSCustomObject] @{
            Type   = 'Footer'
            Output = & $HTMLContent
        }
    }
}
function New-HTMLGage {
    [CmdletBinding()]
    param (
        [scriptblock] $GageContent,
        [validateSet('Gage', 'Donut')][string] $Type = 'Gage',
        [string] $BackgroundGaugageColor,
        [parameter(Mandatory)][decimal] $Value,
        [string] $ValueSymbol,
        [string] $ValueColor,
        [string] $ValueFont,
        [nullable[int]] $MinValue,
        [string] $MinText,
        [nullable[int]] $MaxValue,
        [string] $MaxText,
        [switch] $Reverse,
        [int] $DecimalNumbers,
        [decimal] $GaugageWidth,
        [string] $Title,
        [string] $Label,
        [string] $LabelColor,
        [switch] $Counter,
        [switch] $ShowInnerShadow,
        [switch] $NoGradient,
        [nullable[decimal]] $ShadowOpacity,
        [nullable[int]] $ShadowSize,
        [nullable[int]] $ShadowVerticalOffset,
        [switch] $Pointer,
        [nullable[int]]  $PointerTopLength,
        [nullable[int]] $PointerBottomLength,
        [nullable[int]] $PointerBottomWidth,
        [string] $StrokeColor,
        #[validateSet('none')][string] $PointerStroke,
        [nullable[int]] $PointerStrokeWidth,
        [validateSet('none', 'square', 'round')] $PointerStrokeLinecap,
        [string] $PointerColor,
        [switch] $HideValue,
        [switch] $HideMinMax,
        [switch] $FormatNumber,
        [switch] $DisplayRemaining,
        [switch] $HumanFriendly,
        [int] $HumanFriendlyDecimal,
        [string[]] $SectorColors
    )
    # Make sure JustGage JS is added to source
    $Script:HTMLSchema.Features.JustGage = $true

    # Build Options
    [string] $ID = "Gage" + (Get-RandomStringName -Size 8)
    $Gage = [ordered] @{
        id    = $ID
        value = $Value
    }

    # When null it will be removed as part of cleanup Remove-EmptyValues
    $Gage.shadowSize = $ShadowSize
    $Gage.shadowOpacity = $ShadowOpacity
    $Gage.shadowVerticalOffset = $ShadowVerticalOffset

    if ($DecimalNumbers) {
        $Gage.decimals = $DecimalNumbers
    }
    if ($Title) {
        $Gage.title = $Title
    }
    if ($ValueColor) {
        $Gage.valueFontColor = $ValueColor
    }
    if ($ValueColor) {
        $Gage.valueFontFamily = $ValueFont
    }
    if ($MinText) {
        $Gage.minText = $MinText
    }
    if ($MaxText) {
        $Gage.maxText = $MaxText
    }

    $Gage.min = $MinValue
    $Gage.max = $MaxValue

    if ($Label) {
        $Gage.label = $Label
    }
    if ($LabelColor) {
        $Gage.labelFontColor = ConvertFrom-Color -Color $LabelColor
    }
    if ($Reverse) {
        $Gage.reverse = $Reverse.IsPresent
    }
    if ($Type -eq 'Donut') {
        $Gage.donut = $true
    }
    if ($GaugageWidth) {
        $Gage.gaugageWidthScale = $GaugageWidthScale
    }
    if ($Counter) {
        $Gage.counter = $Counter.IsPresent
    }
    if ($showInnerShadow) {
        $Gage.showInnerShadow = $ShowInnerShadow.IsPresent
    }
    if ($BackgroundGaugageColor) {
        $Gage.gaugeColor = ConvertFrom-Color -Color $BackgroundGaugageColor
    }
    if ($NoGradient) {
        $Gage.noGradient = $NoGradient.IsPresent
    }

    if ($HideMinMax) {
        $Gage.hideMinMax = $HideMinMax.IsPresent
    }
    if ($HideValue) {
        $Gage.hideValue = $HideValue.IsPresent
    }
    if ($FormatNumber) {
        $Gage.formatNumber = $FormatNumber.IsPresent
    }
    if ($DisplayRemaining) {
        $Gage.displayRemaining = $DisplayRemaining.IsPresent
    }
    if ($HumanFriendly) {
        $Gage.humanFriendly = $HumanFriendly.IsPresent
        if ($HumanFriendlyDecimal) {
            $Gage.humanFriendlyDecimal = $HumanFriendlyDecimal
        }
    }
    if ($ValueSymbol) {
        $Gage.symbol = $ValueSymbol
    }

    if ($GageContent) {
        [Array] $GageOutput = & $GageContent
        if ($GageOutput.Count -gt 0) {
            $Gage.customSectors = @{
                percents = $true
                ranges   = $GageOutput
            }
        }
    }



    if ($Pointer) {
        $Gage.pointer = $Pointer.IsPresent

        $Gage.pointerOptions = @{ }
        #if ($PointerToplength) {
        $Gage.pointerOptions.toplength = $PointerTopLength
        #}
        #if ($PointerBottomLength) {
        $Gage.pointerOptions.bottomlength = $PointerBottomLength
        #}
        #if ($PointerBottomWidth) {
        $Gage.pointerOptions.bottomwidth = $PointerBottomWidth
        #}
        #if ($PointerStroke) {

        #}
        #if ($PointerStrokeWidth) {
        $Gage.pointerOptions.stroke_width = $PointerStrokeWidth
        #}
        #if ($PointerStrokeLinecap) {
        $Gage.pointerOptions.stroke_linecap = $PointerStrokeLinecap
        #}
        #if ($PointerColor) {
        $Gage.pointerOptions.color = ConvertFrom-Color -Color $PointerColor
        #}
        #if ($StrokeColor) {
        $Gage.pointerOptions.stroke = ConvertFrom-Color -Color $StrokeColor
        #}
    }
    $gage.relativeGaugeSize = $true
    Remove-EmptyValues -Hashtable $Gage -Rerun 1 -Recursive

    # Build HTML
    $Div = New-HTMLTag -Tag 'div' -Attributes @{ id = $Gage.id; }

    $Script = New-HTMLTag -Tag 'script' -Value {
        # Convert Dictionary to JSON and return chart within SCRIPT tag
        # Make sure to return with additional empty string
        $JSON = $Gage | ConvertTo-Json -Depth 5 | ForEach-Object { [System.Text.RegularExpressions.Regex]::Unescape($_) }
        "document.addEventListener(`"DOMContentLoaded`", function (event) {"
        "var g1 = new JustGage( $JSON );"
        "});"
    } -NewLine

    # Return Data
    $Div
    $Script
}
Register-ArgumentCompleter -CommandName New-HTMLGage -ParameterName GaugageColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLGage -ParameterName LabelColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLGage -ParameterName ValueColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLGage -ParameterName PointerColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLGage -ParameterName StrokeColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLGage -ParameterName SectorColors -ScriptBlock { $Script:RGBColors.Keys }

<#
| | Name                 | Default                           | Description                                                                         |
|-| -------------------- | --------------------------------- | ----------------------------------------------------------------------------------- |
|+| id                   | (required)                        | The HTML container element id                                                       |
|+| value                | 0                                 | Value Gauge is showing                                                              |
| | parentNode           | null                              | The HTML container element object. Used if id is not present                        |
| | defaults             | false                             | Defaults parameters to use globally for gauge objects                               |
| | width                | null                              | The Gauge width in pixels (Integer)                                                 |
| | height               | null                              | The Gauge height in pixels                                                          |
|+| valueFontColor       | #010101                           | Color of label showing current value                                                |
|+| valueFontFamily      | Arial                             | Font of label showing current value                                                 |
|+| symbol               | ''                                | Special symbol to show next to value                                                |
|+| min                  | 0                                 | Min value                                                                           |
|+| minTxt               | false                             | Min value text, overrides min if specified                                          |
|+| max                  | 100                               | Max value                                                                           |
|+| maxTxt               | false                             | Max value text, overrides max if specified                                          |
|+| reverse              | false                             | Reverse min and max                                                                 |
|+| humanFriendlyDecimal | 0                                 | Number of decimal places for our human friendly number to contain                   |
| | textRenderer         | null                              | Function applied before redering text (value) => value                              |
| | onAnimationEnd       | null                              | Function applied after animation is done                                            |
|+| gaugeWidthScale      | 1.0                               | Width of the gauge element                                                          |
|+| gaugeColor           | #edebeb                           | Background color of gauge element                                                   |
|+| label                | ''                                | Text to show below value                                                            |
|+| labelFontColor       | #b3b3b3                           | Color of label showing label under value                                            |
|+| shadowOpacity        | 0.2                               | Shadow opacity 0 ~ 1                                                                |
|+| shadowSize           | 5                                 | Inner shadow size                                                                   |
|+| shadowVerticalOffset | 3                                 | How much shadow is offset from top                                                  |
|+| levelColors          | ["#a9d70b", "#f9c802", "#ff0000"] | Colors of indicator, from lower to upper, in RGB format                             |
| | startAnimationTime   | 700                               | Length of initial animation in milliseconds                                         |
| | startAnimationType   | >                                 | Type of initial animation (linear, >, <, <>, bounce)                                |
| | refreshAnimationTime | 700                               | Length of refresh animation in milliseconds                                         |
| | refreshAnimationType | >                                 | Type of refresh animation (linear, >, <, <>, bounce)                                |
| | donutStartAngle      | 90                                | Angle to start from when in donut mode                                              |
| | valueMinFontSize     | 16                                | Absolute minimum font size for the value label                                      |
| | labelMinFontSize     | 10                                | Absolute minimum font size for the label                                            |
| | minLabelMinFontSize  | 10                                | Absolute minimum font size for the min label                                        |
| | maxLabelMinFontSize  | 10                                | Absolute minimum font size for the man label                                        |
|+| hideValue            | false                             | Hide value text                                                                     |
|+| hideMinMax           | false                             | Hide min/max text                                                                   |
|+| showInnerShadow      | false                             | Show inner shadow                                                                   |
|+| humanFriendly        | false                             | convert large numbers for min, max, value to human friendly (e.g. 1234567 -> 1.23M) |
|+| noGradient           | false                             | Whether to use gradual color change for value, or sector-based                      |
|+| donut                | false                             | Show donut gauge                                                                    |
|*| relativeGaugeSize    | false                             | Whether gauge size should follow changes in container element size                  |
|+| counter              | false                             | Animate text value number change                                                    |
|+| decimals             | 0                                 | Number of digits after floating point                                               |
| | customSectors        | {}                                | Custom sectors colors. Expects an object                                            |
|+| formatNumber         | false                             | Formats numbers with commas where appropriate                                       |
|x| pointer              | false                             | Show value pointer                                                                  |
|x| pointerOptions       | {}                                | Pointer options. Expects an object                                                  |
|+| displayRemaining     | false                             | Replace display number with the value remaining to reach max value                  |
#>

<#
pointerOptions: {
  toplength: null,
  bottomlength: null,
  bottomwidth: null,
  stroke: 'none',
  stroke_width: 0,
  stroke_linecap: 'square',
  color: '#000000'
}
#>
<#
customSectors: {
  percents: true, // lo and hi values are in %
  ranges: [{
    color : "#43bf58",
    lo : 0,
    hi : 50
  },
  {
    color : "#ff3b30",
    lo : 51,
    hi : 100
  }]
}
#>
function New-HTMLHeader {
    [alias('Header')]
    [CmdletBinding()]
    param(
        [scriptblock] $HTMLContent
    )

    if ($HTMLContent) {
        [PSCustomObject] @{
            Type   = 'Header'
            Output = & $HTMLContent
        }
    }
}
Function New-HTMLHeading {
    [CmdletBinding()]
    Param (
        [validateset('h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'h7')][string]$Heading,
        [string]$HeadingText,
        # [validateset('default', 'central')][string] $Type = 'default',
        [switch] $Underline,
        [string]$Color
    )
    if ($null -ne $Color) {
        $RGBcolor = ConvertFrom-Color -Color $Color
        $Attributes = @{
            style = "color: $RGBcolor;"
        }
    } else {
        $Attributes = @{ }
    }
    # if ($Type -eq 'central') {
    #        $Attributes.Class = 'central'
    #   }
    if ($Underline) {
        $Attributes.Class = "$($Attributes.Class) underline"
    }

    New-HTMLTag -Tag $Heading -Attributes $Attributes {
        $HeadingText
    }
}
function New-HTMLHierarchicalTree {
    param(
        [ScriptBlock] $TreeView
    )

    $Script:HTMLSchema.Features.D3Mitch = $true

    [string] $ID = "HierarchicalTree-" + (Get-RandomStringName -Size 8)

    $TreeNodes = [System.Collections.Generic.List[System.Collections.IDictionary]]::new()

    [Array] $Settings = & $TreeView
    foreach ($Object in $Settings) {
        if ($Object.Type -eq 'TreeNode') {
            $TreeNodes.Add($Object.Settings)
        }
    }

    # Prepare NODES
    $Data = $TreeNodes | ConvertTo-Json -Depth 5

    # Prepare HTML
    $Section = New-HTMLTag -Tag 'section' -Attributes @{ id = $ID; class = 'hierarchicalTree' }
    $Script = New-HTMLTag -Tag 'script' -Value { @"
        var data = $Data;
        var treePlugin = new d3.mitchTree.boxedTree()
        .setIsFlatData(true)
        .setData(data)
        .setElement(document.getElementById("$ID"))
        .setIdAccessor(function (data) {
            return data.id;
        })
        .setParentIdAccessor(function (data) {
            return data.parentId;
        })
        .setBodyDisplayTextAccessor(function (data) {
            return data.description;
        })
        .setTitleDisplayTextAccessor(function (data) {
            return data.name;
        })
        .initialize();
"@
    } -NewLine

    # Send HTML
    $Section
    $Script
}
function New-HTMLHorizontalLine {
    [CmdletBinding()]
    param()
    New-HTMLTag -Tag 'hr' -SelfClosing
}
function New-HTMLImage {
    <#
    .SYNOPSIS
    Short description

    .DESCRIPTION
    Long description

    .PARAMETER Source
    Parameter description

    .PARAMETER UrlLink
    Parameter description

    .PARAMETER AlternativeText
    Parameter description

    .PARAMETER Class
    Parameter description

    .PARAMETER Target
    Parameter description

    .PARAMETER Width
    Parameter description

    .PARAMETER Height
    Parameter description

    .EXAMPLE
    New-HTMLImage -Source 'https://evotec.pl/image.png' -UrlLink 'https://evotec.pl/' -AlternativeText 'My other text' -Class 'otehr' -Width '100%'

    .NOTES
    General notes
    #>
    [alias('Image')]
    [CmdletBinding()]
    param(
        [string] $Source,
        [Uri] $UrlLink = '',
        [string] $AlternativeText = '',
        [string] $Class = 'Logo',
        [string] $Target = '_blank',
        [string] $Width,
        [string] $Height
    )

    New-HTMLTag -Tag 'div' -Attributes @{ class = $Class.ToLower() } {
        $AAttributes = [ordered]@{
            'target' = $Target
            'href'   = $UrlLink
        }
        New-HTMLTag -Tag 'a' -Attributes $AAttributes {
            $ImgAttributes = [ordered]@{
                'src'    = "$Source"
                'alt'    = "$AlternativeText"
                'width'  = "$Height"
                'height' = "$Width"
            }
            New-HTMLTag -Tag 'img' -Attributes $ImgAttributes
        }
    }
}
function New-HTMLList {
    [alias('EmailList')]
    [CmdletBinding()]
    param(
        [ScriptBlock]$ListItems,
        [ValidateSet('Unordered', 'Ordered')] [string] $Type = 'Unordered',
        [string] $Color,
        [string] $BackGroundColor,
        [int] $FontSize,
        [ValidateSet('normal', 'bold', 'bolder', 'lighter', '100', '200', '300', '400', '500', '600', '700', '800', '900')][string] $FontWeight,
        [ValidateSet('normal', 'italic', 'oblique')][string] $FontStyle,
        [ValidateSet('normal', 'small-caps')][string] $FontVariant,
        [string] $FontFamily,
        [ValidateSet('left', 'center', 'right', 'justify')][string] $Alignment,
        [ValidateSet('none', 'line-through', 'overline', 'underline')][string] $TextDecoration,
        [ValidateSet('uppercase', 'lowercase', 'capitalize')][string] $TextTransform,
        [ValidateSet('rtl')][string] $Direction,
        [switch] $LineBreak
    )

    $newHTMLSplat = @{ }
    if ($Alignment) {
        $newHTMLSplat.Alignment = $Alignment
    }
    if ($FontSize) {
        $newHTMLSplat.FontSize = $FontSize
    }
    if ($TextTransform) {
        $newHTMLSplat.TextTransform = $TextTransform
    }
    if ($Color) {
        $newHTMLSplat.Color = $Color
    }
    if ($FontFamily) {
        $newHTMLSplat.FontFamily = $FontFamily
    }
    if ($Direction) {
        $newHTMLSplat.Direction = $Direction
    }
    if ($FontStyle) {
        $newHTMLSplat.FontStyle = $FontStyle
    }
    if ($TextDecoration) {
        $newHTMLSplat.TextDecoration = $TextDecoration
    }
    if ($BackGroundColor) {
        $newHTMLSplat.BackGroundColor = $BackGroundColor
    }
    if ($FontVariant) {
        $newHTMLSplat.FontVariant = $FontVariant
    }
    if ($FontWeight) {
        $newHTMLSplat.FontWeight = $FontWeight
    }
    if ($LineBreak) {
        $newHTMLSplat.LineBreak = $LineBreak
    }

    [bool] $SpanRequired = $false
    foreach ($Entry in $newHTMLSplat.GetEnumerator()) {
        if (($Entry.Value | Measure-Object).Count -gt 0) {
            $SpanRequired = $true
            break
        }
    }

    if ($SpanRequired) {
        New-HTMLSpanStyle @newHTMLSplat {
            if ($Type -eq 'Unordered') {
                New-HTMLTag -Tag 'ul' {
                    Invoke-Command -ScriptBlock $ListItems
                }
            } else {
                New-HTMLTag -Tag 'ol' {
                    Invoke-Command -ScriptBlock $ListItems
                }
            }
        }
    } else {
        if ($Type -eq 'Unordered') {
            New-HTMLTag -Tag 'ul' {
                Invoke-Command -ScriptBlock $ListItems
            }
        } else {
            New-HTMLTag -Tag 'ol' {
                Invoke-Command -ScriptBlock $ListItems
            }
        }
    }
}

Register-ArgumentCompleter -CommandName New-HTMLList -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLList -ParameterName BackGroundColor -ScriptBlock { $Script:RGBColors.Keys }
function New-HTMLListItem {
    [CmdletBinding()]
    param(
        [string[]] $Text,
        [string[]] $Color = @(),
        [string[]] $BackGroundColor = @(),
        [int[]] $FontSize = @(),
        [ValidateSet('normal', 'bold', 'bolder', 'lighter', '100', '200', '300', '400', '500', '600', '700', '800', '900')][string[]] $FontWeight = @(),
        [ValidateSet('normal', 'italic', 'oblique')][string[]] $FontStyle = @(),
        [ValidateSet('normal', 'small-caps')][string[]] $FontVariant = @(),
        [string[]] $FontFamily = @(),
        [ValidateSet('left', 'center', 'right', 'justify')][string[]] $Alignment = @(),
        [ValidateSet('none', 'line-through', 'overline', 'underline')][string[]] $TextDecoration = @(),
        [ValidateSet('uppercase', 'lowercase', 'capitalize')][string[]] $TextTransform = @(),
        [ValidateSet('rtl')][string[]] $Direction = @(),
        [switch] $LineBreak
    )

    $newHTMLTextSplat = @{
        Alignment       = $Alignment
        FontSize        = $FontSize
        TextTransform   = $TextTransform
        Text            = $Text
        Color           = $Color
        FontFamily      = $FontFamily
        Direction       = $Direction
        FontStyle       = $FontStyle
        TextDecoration  = $TextDecoration
        BackGroundColor = $BackGroundColor
        FontVariant     = $FontVariant
        FontWeight      = $FontWeight
        LineBreak       = $LineBreak
    }

    if (($FontSize.Count -eq 0) -or ($FontSize -eq 0)) {
        $Size = ''
    } else {
        $size = "$($FontSize)px"
    }
    $Style = @{
        style = @{
            'color'            = ConvertFrom-Color -Color $Color
            'background-color' = ConvertFrom-Color -Color $BackGroundColor
            'font-size'        = $Size
            'font-weight'      = $FontWeight
            'font-variant'     = $FontVariant
            'font-family'      = $FontFamily
            'font-style'       = $FontStyle
            'text-align'       = $Alignment


            'text-decoration'  = $TextDecoration
            'text-transform'   = $TextTransform
            'direction'        = $Direction
        }
    }

    New-HTMLTag -Tag 'li' -Attributes $Style -Value {
        New-HTMLText @newHTMLTextSplat -SkipParagraph
    }
}

Register-ArgumentCompleter -CommandName New-HTMLListItem -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLListItem -ParameterName BackgroundColor -ScriptBlock { $Script:RGBColors.Keys }
function New-HTMLLogo {
    [CmdletBinding()]
    param(
        [String] $LogoPath,
        [string] $LeftLogoName = "Sample",
        [string] $RightLogoName = "Alternate",
        [string] $LeftLogoString,
        [string] $RightLogoString,
        [switch] $HideLogos
    )

    $LogoSources = Get-HTMLLogos `
        -RightLogoName $RightLogoName `
        -LeftLogoName $LeftLogoName  `
        -LeftLogoString $LeftLogoString `
        -RightLogoString $RightLogoString

    #Convert-StyleContent1 -Options $Options

    $Options = [PSCustomObject] @{
        Logos        = $LogoSources
        ColorSchemes = $ColorSchemes
    }

    if ($HideLogos -eq $false) {
        $Leftlogo = $Options.Logos[$LeftLogoName]
        $Rightlogo = $Options.Logos[$RightLogoName]
        $Script:HTMLSchema.Logos = @(
            '<!-- START LOGO -->'
            New-HTMLTag -Tag 'div' -Attributes @{ class = 'logos' } {
                New-HTMLTag -Tag 'div' -Attributes @{ class = 'leftLogo' } {
                    New-HTMLTag -Tag 'img' -Attributes @{ src = "$LeftLogo" } -SelfClosing
                }
                New-HTMLTag -Tag 'div' -Attributes @{ class = 'rightLogo' } {
                    New-HTMLTag -Tag 'img' -Attributes @{ src = "$RightLogo" } -SelfClosing
                }
            }
            '<!-- END LOGO -->'
        ) -join ''
    }
}
function New-HTMLMain {
    [alias('Main')]
    [CmdletBinding()]
    param(
        [scriptblock] $HTMLContent
    )
    if ($HTMLContent) {
        [PSCustomObject] @{
            Type   = 'Main'
            Output = & $HTMLContent
        }
    }
}
Function New-HTMLPanel {
    [alias('New-HTMLColumn', 'Panel')]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, Position = 0)][ValidateNotNull()][ScriptBlock] $Content = $(Throw "Open curly brace with Content"),
        #[alias('ColumnCount', 'Columns')][ValidateSet('1', '2', '3', '4', '5 ', '6', '7', '8', '9', '10', '11', '12')][string] $Count = 1,
        [alias('BackgroundShade')][string]$BackgroundColor,
        [switch] $Invisible,
        [alias('flex-basis')][string] $Width,
        [string] $Margin #,
        # [int] $Height
    )
    #if ($Height -ne 0) {
    #     $StyleHeight = "height: $($Height)px"
    #}
    # $StyleWidth = "width: calc(100% / $Count - 10px)"

    if ($BackgroundColor) {
        $BackGroundColorFromRGB = ConvertFrom-Color -Color $BackgroundColor
        $DivColumnStyle = "background-color:$BackGroundColorFromRGB;"
    } else {
        $DivColumnStyle = ""
    }
    if ($Invisible) {
        $DivColumnStyle = "$DivColumnStyle box-shadow: unset !important;"
    }

    if ($Width -or $Margin) {
        [string] $ClassName = "flexPanel$(Get-RandomStringName -Size 8 -LettersOnly)"
        $Attributes = @{
            'flex-basis' = if ($Width) { $Width } else { '100%' }
            'margin'     = if ($Margin) { $Margin }
        }
        $Css = ConvertTo-CSS -ClassName $ClassName -Attributes $Attributes

        $Script:HTMLSchema.CustomCSS.Add($Css)
        [string] $Class = "$ClassName overflowHidden"
    } else {
        [string] $Class = 'flexPanel overflowHidden'
    }

    # New-HTMLTag -Tag 'div' -Attributes @{ class = "flexPanel roundedCorners"; style = $DivColumnStyle } {
    New-HTMLTag -Tag 'div' -Attributes @{ class = "$Class roundedCorners overflowHidden"; style = $DivColumnStyle } {
        Invoke-Command -ScriptBlock $Content
    }
}

Register-ArgumentCompleter -CommandName New-HTMLPanel -ParameterName BackgroundColor -ScriptBlock { $Script:RGBColors.Keys }
Function New-HTMLSection {
    [alias('New-HTMLContent', 'Section')]
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $false, Position = 0)][ValidateNotNull()][ScriptBlock] $Content = $(Throw "Open curly brace"),
        [alias('Name')][Parameter(Mandatory = $false)][string]$HeaderText,
        [alias('TextColor')][string]$HeaderTextColor = "White",
        [alias('TextAlignment')][string][ValidateSet('center', 'left', 'right', 'justify')] $HeaderTextAlignment = 'center',
        [alias('TextBackGroundColor')][string]$HeaderBackGroundColor = "DeepSkyBlue",
        [alias('BackgroundShade')][string]$BackgroundColor = "",
        [alias('Collapsable')][Parameter(Mandatory = $false)][switch] $CanCollapse,
        [Parameter(Mandatory = $false)][switch] $IsHidden,
        [switch] $Collapsed,
        [int] $Height,
        [switch] $Invisible,
        # Following are based on https://css-tricks.com/snippets/css/a-guide-to-flexbox/
        [string][ValidateSet('wrap', 'nowrap', 'wrap-reverse')] $Wrap,
        [string][ValidateSet('row', 'row-reverse', 'column', 'column-reverse')] $Direction,
        [string][ValidateSet('flex-start', 'flex-end', 'center', 'space-between', 'space-around', 'stretch')] $AlignContent,
        [string][ValidateSet('stretch', 'flex-start', 'flex-end', 'center', 'baseline')] $AlignItems,
        [string][ValidateSet('flex-start', 'flex-end', 'center')] $JustifyContent

    )
    $RandomNumber = Get-Random
    $TextHeaderColorFromRGB = ConvertFrom-Color -Color $HeaderTextColor

    $HiddenDivStyle = @{ }

    if ($CanCollapse) {
        $Script:HTMLSchema.Features.HideSection = $true
        if ($IsHidden) {
            $ShowStyle = "color: $TextHeaderColorFromRGB;" # shows Show button
            $HideStyle = "color: $TextHeaderColorFromRGB; display:none;" # hides Hide button
        } else {
            if ($Collapsed) {
                $HideStyle = "color: $TextHeaderColorFromRGB; display:none;" # hides Show button
                $ShowStyle = "color: $TextHeaderColorFromRGB;" # shows Hide button
                # $HiddenDivStyle = 'display:none; '
                $HiddenDivStyle['display'] = 'none'
            } else {
                $ShowStyle = "color: $TextHeaderColorFromRGB; display:none;" # hides Show button
                $HideStyle = "color: $TextHeaderColorFromRGB;" # shows Hide button
            }
        }
    } else {
        if ($IsHidden) {
            $ShowStyle = "color: $TextHeaderColorFromRGB;" # shows Show button
            $HideStyle = "color: $TextHeaderColorFromRGB; display:none;" # hides Hide button
        } else {
            $ShowStyle = "color: $TextHeaderColorFromRGB; display:none;" # hides Show button
            $HideStyle = "color: $TextHeaderColorFromRGB; display:none;" # hides Show button
        }
    }
    if ($IsHidden) {
        $DivContentStyle = @{
            "display"          = 'none'
            #"width"            = "calc(100% / $Count - 15px)"
            #"height"           = if ($Height -ne 0) { "$($Height)px" } else { '' }
            "background-color" = ConvertFrom-Color -Color $BackgroundColor
        }
    } else {
        $DivContentStyle = @{
            # "width"            = "calc(100% / $Count - 15px)"
            #"height"           = if ($Height -ne 0) { "$($Height)px" } else { '' }
            "background-color" = ConvertFrom-Color -Color $BackgroundColor
        }
    }

    $HiddenDivStyle['height'] = if ($Height -ne 0) { "$($Height)px" } else { '' }

    <#
    .flexParent {
        display: flex;
        flex-wrap: nowrap;
        justify-content: space-between;
        padding: 2px;
        /*
        overflow: hidden;
        overflow-x: hidden;
        overflow-y: hidden;
        */
    }
    #>

    if ($Wrap -or $Direction) {
        [string] $ClassName = "flexParent$(Get-RandomStringName -Size 8 -LettersOnly)"
        $Attributes = @{
            'display'        = 'flex'
            'flex-wrap'      = if ($Wrap) { $Wrap } else { }
            'flex-direction' = if ($Direction) { $Direction } else { }
            'align-content'  = if ($AlignContent) { $AlignContent } else { }
            'align-items'    = if ($AlignItems) { $AlignItems } else { }
        }
        $Css = ConvertTo-CSS -ClassName $ClassName -Attributes $Attributes

        $Script:HTMLSchema.CustomCSS.Add($Css)
    } else {
        [string] $ClassName = 'flexParent flexElement overflowHidden'
    }

    $DivHeaderStyle = @{
        "text-align"       = $HeaderTextAlignment
        "background-color" = ConvertFrom-Color -Color $HeaderBackGroundColor
    }
    $HeaderStyle = "color: $TextHeaderColorFromRGB;"
    if ($Invisible) {
        #New-HTMLTag -Tag 'div' -Attributes @{ class = 'flexParentInvisible' } -Value {
        New-HTMLTag -Tag 'div' -Attributes @{ class = $ClassName } -Value {
            New-HTMLTag -Tag 'div' -Attributes @{ class = $ClassName; Style = @{'justify-content' = $JustifyContent } } -Value {
                # New-HTMLTag -Tag 'div' -Attributes @{ class = 'flexParentInvisible flexElement' } -Value {
                $Object = Invoke-Command -ScriptBlock $Content
                if ($null -ne $Object) {
                    $Object
                }
            }
        }
    } else {
        # return this HTML
        New-HTMLTag -Tag 'div' -Attributes @{ 'class' = "defaultSection roundedCorners overflowHidden"; 'style' = $DivContentStyle } -Value {
            New-HTMLTag -Tag 'div' -Attributes @{ 'class' = "defaultHeader"; 'style' = $DivHeaderStyle } -Value {
                New-HTMLAnchor -Name $HeaderText -Text "$HeaderText " -Style $HeaderStyle
                New-HTMLAnchor -Id "show_$RandomNumber" -Href 'javascript:void(0)' -OnClick "show('$RandomNumber');" -Style $ShowStyle -Text '(Show)'
                New-HTMLAnchor -Id "hide_$RandomNumber" -Href 'javascript:void(0)' -OnClick "hide('$RandomNumber');" -Style $HideStyle -Text '(Hide)'
            }
            New-HTMLTag -Tag 'div' -Attributes @{ class = $ClassName; id = $RandomNumber; Style = $HiddenDivStyle } -Value {
                New-HTMLTag -Tag 'div' -Attributes @{ class = "$ClassName collapsable"; id = $RandomNumber; Style = @{'justify-content' = $JustifyContent } } -Value {
                    $Object = Invoke-Command -ScriptBlock $Content
                    if ($null -ne $Object) {
                        $Object
                    }
                }
            }
        }
    }
}

Register-ArgumentCompleter -CommandName New-HTMLSection -ParameterName HeaderTextColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLSection -ParameterName HeaderBackGroundColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLSection -ParameterName BackgroundColor -ScriptBlock { $Script:RGBColors.Keys }

function New-HTMLSpanStyle {
    [CmdletBinding()]
    param(
        [ScriptBlock] $Content,
        [string] $Color,
        [string] $BackGroundColor,
        [int] $FontSize,
        [ValidateSet('normal', 'bold', 'bolder', 'lighter', '100', '200', '300', '400', '500', '600', '700', '800', '900')][string] $FontWeight,
        [ValidateSet('normal', 'italic', 'oblique')][string] $FontStyle,
        [ValidateSet('normal', 'small-caps')][string] $FontVariant,
        [string] $FontFamily,
        [ValidateSet('left', 'center', 'right', 'justify')][string]  $Alignment,
        [ValidateSet('none', 'line-through', 'overline', 'underline')][string]  $TextDecoration,
        [ValidateSet('uppercase', 'lowercase', 'capitalize')][string]  $TextTransform,
        [ValidateSet('rtl')][string] $Direction,
        [switch] $LineBreak
    )
    if ($FontSize -eq 0) {
        $Size = ''
    } else {
        $size = "$($FontSize)px"
    }
    $Style = @{
        style = @{
            'color'            = ConvertFrom-Color -Color $Color
            'background-color' = ConvertFrom-Color -Color $BackGroundColor
            'font-size'        = $Size
            'font-weight'      = $FontWeight
            'font-variant'     = $FontVariant
            'font-family'      = $FontFamily
            'font-style'       = $FontStyle
            'text-align'       = $Alignment


            'text-decoration'  = $TextDecoration
            'text-transform'   = $TextTransform
            'direction'        = $Direction
        }
    }

    if ($Alignment) {
        $StyleDiv = @{ }
        $StyleDiv.Align = $Alignment

        New-HTMLTag -Tag 'div' -Attributes $StyleDiv {
            New-HTMLTag -Tag 'span' -Attributes $Style {
                Invoke-Command -ScriptBlock $Content
            }
        }
    } else {
        New-HTMLTag -Tag 'span' -Attributes $Style {
            Invoke-Command -ScriptBlock $Content
        }
    }
}

Register-ArgumentCompleter -CommandName New-HTMLSpanStyle -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLSpanStyle -ParameterName BackGroundColor -ScriptBlock { $Script:RGBColors.Keys }
function New-HTMLStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)][alias('')][ScriptBlock] $Content
    )
    $Script:HTMLSchema.Features.StatusButtonical = $true
    New-HTMLTag -Tag 'div' -Attributes @{ class = 'buttonicalService' } {
        #New-HTMLTag -Tag 'div' -Attributes @{ class = 'buttonical-align' } {
        Invoke-Command -ScriptBlock $Content
        # }
    }

}
function New-HTMLStatusItem {
    [CmdletBinding()]
    param(
        [string] $ServiceName,
        [string] $ServiceStatus,
        [ValidateSet('Dead', 'Bad', 'Good')]$Icon = 'Good',
        [ValidateSet('0%', '10%', '30%', '70%', '100%')][string] $Percentage = '100%'
    )
    #$Script:HTMLSchema.Features.StatusButtonical = $true
    if ($Icon -eq 'Dead') {
        $IconType = 'performanceDead'
    } elseif ($Icon -eq 'Bad') {
        $IconType = 'performanceProblem'
    } elseif ($Icon -eq 'Good') {
        #$IconType = 'performance'
    }

    if ($Percentage -eq '100%') {
        $Colors = 'background-color: #0ef49b;'
    } elseif ($Percentage -eq '70%') {
        $Colors = 'background-color: #d2dc69;'
    } elseif ($Percentage -eq '30%') {
        $Colors = 'background-color: #faa04b;'
    } elseif ($Percentage -eq '10%') {
        $Colors = 'background-color: #ff9035;'
    } elseif ($Percentage -eq '0%') {
        $Colors = 'background-color: #ff5a64;'
    }

    New-HTMLTag -Tag 'div' -Attributes @{ class = 'buttonical'; style = $Colors } -Value {
        New-HTMLTag -Tag 'div' -Attributes @{ class = 'label' } {
            New-HTMLTag -Tag 'span' -Attributes @{ class = 'performance' } {
                $ServiceName
            }
        }
        New-HTMLTag -Tag 'div' -Attributes @{ class = 'middle' }
        New-HTMLTag -Tag 'div' -Attributes @{ class = 'status' } {
            New-HTMLTag -Tag 'input' -Attributes @{ name = Get-Random; type = 'radio'; value = 'other-item'; checked = 'true' } -SelfClosing
            New-HTMLTag -Tag 'span' -Attributes @{ class = "performance $IconType" } {
                $ServiceStatus
            }
        }
    }
}
function New-HTMLTab {
    [alias('Tab')]
    [CmdLetBinding(DefaultParameterSetName = 'FontAwesomeBrands')]
    param(
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [Parameter(Mandatory = $false, Position = 0)][ValidateNotNull()][ScriptBlock] $HtmlData = $(Throw "No curly brace?)"),
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [alias('TabHeading')][Parameter(Mandatory = $false, Position = 1)][String]$Heading,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")]
        [alias('TabName')][string] $Name = 'Tab',

        # ICON BRANDS
        [ArgumentCompleter(
            {
                param($CommandName, $ParameterName, $WordToComplete, $CommandAst, $FakeBoundParameters)
                ($Global:HTMLIcons.FontAwesomeBrands.Keys)
            }
        )]
        [ValidateScript(
            {
                $_ -in (($Global:HTMLIcons.FontAwesomeBrands.Keys))
            }
        )]
        [parameter(ParameterSetName = "FontAwesomeBrands")][string] $IconBrands,

        # ICON REGULAR
        [ArgumentCompleter(
            {
                param($CommandName, $ParameterName, $WordToComplete, $CommandAst, $FakeBoundParameters)
                ($Global:HTMLIcons.FontAwesomeRegular.Keys)
            }
        )]
        [ValidateScript(
            {
                $_ -in (($Global:HTMLIcons.FontAwesomeRegular.Keys))
            }
        )]
        [parameter(ParameterSetName = "FontAwesomeRegular")][string] $IconRegular,

        # ICON SOLID
        [ArgumentCompleter(
            {
                param($CommandName, $ParameterName, $WordToComplete, $CommandAst, $FakeBoundParameters)
                ($Global:HTMLIcons.FontAwesomeSolid.Keys)
            }
        )]
        [ValidateScript(
            {
                $_ -in (($Global:HTMLIcons.FontAwesomeSolid.Keys))
            }
        )]
        [parameter(ParameterSetName = "FontAwesomeSolid")][string] $IconSolid,

        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")][int] $TextSize,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")][string] $TextColor,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")][int] $IconSize,
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")][string] $IconColor,
        [ValidateSet('uppercase', 'lowercase', 'capitalize')][string] $TextTransform = 'uppercase'  # New-HTMLTab - Add text-transform
    )
    if (-not $Script:HTMLSchema.Features) {
        Write-Warning 'New-HTMLTab - Creation of HTML aborted. Most likely New-HTML is missing.'
        Exit
    }
    [string] $Icon = ''
    if ($IconBrands) {
        $Icon = "fab fa-$IconBrands" # fa-$($FontSize)x"
    } elseif ($IconRegular) {
        $Icon = "far fa-$IconRegular" # fa-$($FontSize)x"
    } elseif ($IconSolid) {
        $Icon = "fas fa-$IconSolid" # fa-$($FontSize)x"
    }

    $StyleText = @{ }
    if ($TextSize -ne 0) {
        $StyleText.'font-size' = "$($TextSize)px"
    }
    if ($TextColor) {
        $StyleText.'color' = ConvertFrom-Color -Color $TextColor
    }
    # New-HTMLTab - Add text-transform
    $StyleText.'text-transform' = "$TextTransform"
    # end

    $StyleIcon = @{ }
    if ($IconSize -ne 0) {
        $StyleIcon.'font-size' = "$($IconSize)px"
    }
    if ($IconColor) {
        $StyleIcon.'color' = ConvertFrom-Color -Color $IconColor
    }
    $Script:HTMLSchema.Features.Tabbis = $true

    # Reset all Tabs Headers to make sure there are no Current Tab Set
    # This is required for New-HTMLTable

    foreach ($Tab in $Script:HTMLSchema.TabsHeaders) {
        $Tab.Current = $false
    }

    # Start Tab Tracking
    $Tab = [ordered] @{ }
    $Tab.ID = "Tab-$(Get-RandomStringName -Size 8)"
    $Tab.Name = " $Name"
    $Tab.StyleIcon = $StyleIcon
    $Tab.StyleText = $StyleText
    #$Tab.Used = $true
    $Tab.Current = $true


    if ($Script:HTMLSchema.TabsHeaders | Where-Object { $_.Active -eq $true }) {
        $Tab.Active = $false
    } else {
        $Tab.Active = $true
    }

    # $Tab.Active = $true
    # $Tab.Active = $true
    $Tab.Icon = $Icon
    # End Tab Tracking

    # This is building HTML

    if ($Tab.Active) {
        $Class = 'active'
    } else {
        $Class = ''
    }
    #New-HTMLTag -Tag 'div' -Attributes @{ id = $Tab.ID; class = $Class } {
    New-HTMLTag -Tag 'div' -Attributes @{ id = $Tab.ID; class = $Class } {
        if (-not [string]::IsNullOrWhiteSpace($Heading)) {
            New-HTMLTag -Tag 'h7' {
                $Heading
            }
        }
        $OutputHTML = Invoke-Command -ScriptBlock $HtmlData
        [Array] $TabsCollection = foreach ($_ in $OutputHTML) {
            if ($_ -is [System.Collections.IDictionary]) {
                $_
                $Script:HTMLSchema.TabsHeadersNested.Add($_)
            }
        }
        [Array] $HTML = foreach ($_ in $OutputHTML) {
            if ($_ -isnot [System.Collections.IDictionary]) {
                $_
            }
        }
        if ($TabsCollection.Count -gt 0) {
            New-HTMLTabHead -TabsCollection $TabsCollection
            New-HTMLTag -Tag 'div' -Attributes @{ 'data-panes' = 'true' } {
                # Add remaining data
                $HTML
            }

        } else {
            $HTML
        }
    }
    $Script:HTMLSchema.TabsHeaders.Add($Tab)
    $Tab
}

Register-ArgumentCompleter -CommandName New-HTMLTab -ParameterName IconColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLTab -ParameterName TextColor -ScriptBlock { $Script:RGBColors.Keys }

function New-HTMLTable {
    [alias('Table', 'EmailTable')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)][ScriptBlock] $HTML,
        [Parameter(Mandatory = $false, Position = 1)][ScriptBlock] $PreContent,
        [Parameter(Mandatory = $false, Position = 2)][ScriptBlock] $PostContent,
        [alias('ArrayOfObjects', 'Object', 'Table')][Array] $DataTable,
        [string[]][ValidateSet('copyHtml5', 'excelHtml5', 'csvHtml5', 'pdfHtml5', 'pageLength')] $Buttons = @('copyHtml5', 'excelHtml5', 'csvHtml5', 'pdfHtml5', 'pageLength'),
        [string[]][ValidateSet('numbers', 'simple', 'simple_numbers', 'full', 'full_numbers', 'first_last_numbers')] $PagingStyle = 'full_numbers',
        [int[]]$PagingOptions = @(15, 25, 50, 100),
        [switch]$DisablePaging,
        [switch]$DisableOrdering,
        [switch]$DisableInfo,
        [switch]$HideFooter,
        [switch]$DisableColumnReorder,
        [switch]$DisableProcessing,
        [switch]$DisableResponsiveTable,
        [switch]$DisableSelect,
        [switch]$DisableStateSave,
        [switch]$DisableSearch,
        [switch]$ScrollCollapse,
        [switch]$OrderMulti,
        [switch]$Filtering,
        [ValidateSet('Top', 'Bottom', 'Both')][string]$FilteringLocation = 'Bottom',
        [string[]][ValidateSet('display', 'cell-border', 'compact', 'hover', 'nowrap', 'order-column', 'row-border', 'stripe')] $Style = @('display', 'compact'),
        [switch]$Simplify,
        [string]$TextWhenNoData = 'No data available.',
        [int] $ScreenSizePercent = 0,
        [string[]] $DefaultSortColumn,
        [int[]] $DefaultSortIndex,
        [ValidateSet('Ascending', 'Descending')][string] $DefaultSortOrder = 'Ascending',
        [string[]] $DateTimeSortingFormat,
        [alias('Search')][string]$Find,
        [switch] $InvokeHTMLTags,
        [switch] $DisableNewLine,
        [switch] $ScrollX,
        [switch] $ScrollY,
        [int] $ScrollSizeY = 500,
        [int] $FreezeColumnsLeft,
        [int] $FreezeColumnsRight,
        [switch] $FixedHeader,
        [switch] $FixedFooter,
        [string[]] $ResponsivePriorityOrder,
        [int[]] $ResponsivePriorityOrderIndex,
        [string[]] $PriorityProperties,
        [alias('DataTableName')][string] $DataTableID,
        [switch] $ImmediatelyShowHiddenDetails,
        [alias('RemoveShowButton')][switch] $HideShowButton,
        [switch] $AllProperties,
        [switch] $SkipProperties,
        [switch] $Compare,
        [alias('CompareWithColors')][switch] $HighlightDifferences,
        [int] $First,
        [int] $Last,
        [alias('Replace')][Array] $CompareReplace
    )
    if (-not $Script:HTMLSchema.Features) {
        Write-Warning 'New-HTMLTable - Creation of HTML aborted. Most likely New-HTML is missing.'
        Exit
    }
    # Theme creator  https://datatables.net/manual/styling/theme-creator
    $ConditionalFormatting = [System.Collections.Generic.List[PSCustomObject]]::new()
    $CustomButtons = [System.Collections.Generic.List[PSCustomObject]]::new()
    $HeaderRows = [System.Collections.Generic.List[PSCustomObject]]::new()
    $HeaderStyle = [System.Collections.Generic.List[PSCustomObject]]::new()
    $HeaderTop = [System.Collections.Generic.List[PSCustomObject]]::new()
    $HeaderResponsiveOperations = [System.Collections.Generic.List[PSCustomObject]]::new()
    $ContentRows = [System.Collections.Generic.List[PSCustomObject]]::new()
    $ContentStyle = [System.Collections.Generic.List[PSCustomObject]]::new()
    $ContentTop = [System.Collections.Generic.List[PSCustomObject]]::new()
    $ContentFormattingInline = [System.Collections.Generic.List[PSCustomObject]]::new()
    $ReplaceCompare = [System.Collections.Generic.List[System.Collections.IDictionary]]::new()
    $RowGrouping = @{ }

    if ($HTML) {
        [Array] $Output = & $HTML

        if ($Output.Count -gt 0) {
            foreach ($Parameters in $Output) {
                if ($Parameters.Type -eq 'TableButtonPDF') {
                    $CustomButtons.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableButtonCSV') {
                    $CustomButtons.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableButtonPageLength') {
                    $CustomButtons.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableButtonExcel') {
                    $CustomButtons.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableButtonPDF') {
                    $CustomButtons.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableButtonPrint') {
                    $CustomButtons.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableButtonCopy') {
                    $CustomButtons.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableCondition') {
                    $ConditionalFormatting.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableHeaderMerge') {
                    $HeaderRows.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableHeaderStyle') {
                    $HeaderStyle.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableHeaderFullRow') {
                    $HeaderTop.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableContentMerge') {
                    $ContentRows.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableContentStyle') {
                    $ContentStyle.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableContentFullRow') {
                    $ContentTop.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableConditionInline') {
                    $ContentFormattingInline.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableHeaderResponsiveOperations') {
                    $HeaderResponsiveOperations.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableReplaceCompare') {
                    $ReplaceCompare.Add($Parameters.Output)
                } elseif ($Parameters.Type -eq 'TableRowGrouping') {
                    $RowGrouping = $Parameters.Output
                }
            }
        }
    }


    # Limit objects count First or Last
    if ($First -or $Last) {
        $DataTable = $DataTable | Select-Object -First $First -Last $Last
    }

    if ($Compare) {
        $Splitter = "`r`n"

        if ($ReplaceCompare) {
            foreach ($R in $CompareReplace) {
                $ReplaceCompare.Add($R)
            }
        }

        $DataTable = Compare-MultipleObjects -Objects $DataTable -Summary -Splitter $Splitter -FormatOutput -AllProperties:$AllProperties -SkipProperties:$SkipProperties -Replace $ReplaceCompare

        if ($HighlightDifferences) {
            $Highlight = for ($i = 0; $i -lt $DataTable.Count; $i++) {
                if ($DataTable[$i].Status -eq $false) {
                    # Different row
                    foreach ($DifferenceColumn in $DataTable[$i].Different) {
                        $DataSame = $DataTable[$i]."$DifferenceColumn-Same" -join $Splitter
                        $DataAdd = $DataTable[$i]."$DifferenceColumn-Add" -join $Splitter
                        $DataRemove = $DataTable[$i]."$DifferenceColumn-Remove" -join $Splitter

                        if ($DataSame -ne '') {
                            $DataSame = "$DataSame$Splitter"
                        }
                        if ($DataAdd -ne '') {
                            $DataAdd = "$DataAdd$Splitter"
                        }
                        if ($DataRemove -ne '') {
                            $DataRemove = "$DataRemove$Splitter"
                        }
                        $Text = New-HTMLText -Text $DataSame, $DataRemove, $DataAdd -Color Black, Red, Blue -TextDecoration none, line-through, none -FontWeight normal, bold, bold
                        New-TableContent -ColumnName "$DifferenceColumn" -RowIndex ($i + 1) -Text "$Text"
                    }
                } else {
                    # Same row
                    # New-TableContent -RowIndex ($i + 1) -BackGroundColor Green -Color White
                }
            }
        }
        $Properties = Select-Properties -Objects $DataTable -ExcludeProperty '*-*', 'Same', 'Different'
        $DataTable = $DataTable | Select-Object -Property $Properties

        if ($HighlightDifferences) {
            foreach ($Parameter in $Highlight.Output) {
                $ContentStyle.Add($Parameter)
            }
        }
    }

    if ($AllProperties) {
        $Properties = Select-Properties -Objects $DataTable -AllProperties:$AllProperties
        $DataTable = $DataTable | Select-Object -Property $Properties
    }

    # This is more direct way of PriorityProperties that will work also on Scroll and in other circumstances
    if ($PriorityProperties) {
        if ($DataTable.Count -gt 0) {
            $Properties = $DataTable[0].PSObject.Properties.Name
            # $Properties = Select-Properties -Objects $DataTable -AllProperties:$AllProperties
            $RemainingProperties = foreach ($Property in $Properties) {
                if ($PriorityProperties -notcontains $Property) {
                    $Property
                }
            }
            $BoundedProperties = $PriorityProperties + $RemainingProperties
            $DataTable = $DataTable | Select-Object -Property $BoundedProperties
        }
    }

    # This option disable paging if number of elements is less or equal count of elements in DataTable
    $PagingOptions = $PagingOptions | Sort-Object -Unique
    if ($DataTable.Count -le $PagingOptions[0]) {
        $DisablePaging = $true
    }

    # Building HTML Table / Script
    if (-not $DataTableID) {
        # Only define this if user failed to deliver as per https://github.com/EvotecIT/PSWriteHTML/issues/29
        $DataTableID = "DT-$(Get-RandomStringName -Size 8 -LettersOnly)" # this builds table ID
    }
    if ($null -eq $DataTable -or $DataTable.Count -eq 0) {
        #return ''
        $Filtering = $false # setting it to false because it's not nessecary
        $HideFooter = $true
        $DataTable = $TextWhenNoData
    }
    if ($DataTable[0] -is [System.Collections.IDictionary]) {
        Write-Verbose 'New-HTMLTable - Working with IDictionary'
        [Array] $TemporaryTable = foreach ($_ in $DataTable) {
            $_.GetEnumerator() | Select-Object Name, Value
        }
        [Array] $Table = $TemporaryTable | ConvertTo-Html -Fragment | Select-Object -SkipLast 1 | Select-Object -Skip 2 # This removes table tags (open/closing)
        #[Array] $Table = $($DataTable).GetEnumerator() | Select-Object Name, Value | ConvertTo-Html -Fragment | Select-Object -SkipLast 1 | Select-Object -Skip 2 # This removes table tags (open/closing)
    } elseif ($DataTable[0] -is [string]) {
        [Array] $Table = $DataTable | ForEach-Object { [PSCustomObject]@{ 'Name' = $_ } } | ConvertTo-Html -Fragment | Select-Object -SkipLast 1 | Select-Object -Skip 2
    } else {
        Write-Verbose 'New-HTMLTable - Working with Objects'
        [Array] $Table = $DataTable | ConvertTo-Html -Fragment | Select-Object -SkipLast 1 | Select-Object -Skip 2 # This removes table tags (open/closing)
    }
    [string] $Header = $Table | Select-Object -First 1 # this gets header
    [string[]] $HeaderNames = $Header -replace '</th></tr>' -replace '<tr><th>' -split '</th><th>'
    $AddedHeader = Add-TableHeader -HeaderRows $HeaderRows -HeaderNames $HeaderNames -HeaderStyle $HeaderStyle -HeaderTop $HeaderTop -HeaderResponsiveOperations $HeaderResponsiveOperations

    # This modifies Table content.
    # It basically goes thru every single row and checks if values to add styles or inline conditional formatting
    # It's heavier then JS, so use when nessecary
    if ($ContentRows.Capacity -gt 0 -or $ContentStyle.Count -gt 0 -or $ContentTop.Count -gt 0 -or $ContentFormattingInline.Count -gt 0) {
        $Table = Add-TableContent -ContentRows $ContentRows -ContentStyle $ContentStyle -ContentTop $ContentTop -ContentFormattingInline $ContentFormattingInline -Table $Table -HeaderNames $HeaderNames
    }


    $Table = $Table | Select-Object -Skip 1 # this gets actuall table content
    $Options = [ordered] @{
        <# DOM Definition: https://datatables.net/reference/option/dom
            l - length changing input control
            f - filtering input
            t - The table!
            i - Table information summary
            p - pagination control
            r - processing display element
            B - Buttons
            S - Select
            F - FadeSeaerch
        #>
        dom              = 'Bfrtip'
        #buttons          = @($Buttons)
        buttons          = @(
            if ($CustomButtons) {
                $CustomButtons
            } else {
                foreach ($button in $Buttons) {
                    if ($button -ne 'pdfHtml5') {
                        @{
                            extend = $button
                        }
                    } else {
                        @{
                            extend      = 'pdfHtml5'
                            pageSize    = 'A3'
                            orientation = 'landscape'
                        }
                    }
                }
            }
        )
        "searchFade"     = $false
        "colReorder"     = -not $DisableColumnReorder.IsPresent


        # https://datatables.net/examples/basic_init/scroll_y_dynamic.html
        "paging"         = -not $DisablePaging
        "scrollCollapse" = $ScrollCollapse.IsPresent

        <# Paging Type
            numbers - Page number buttons only
            simple - 'Previous' and 'Next' buttons only
            simple_numbers - 'Previous' and 'Next' buttons, plus page numbers
            full - 'First', 'Previous', 'Next' and 'Last' buttons
            full_numbers - 'First', 'Previous', 'Next' and 'Last' buttons, plus page numbers
            first_last_numbers - 'First' and 'Last' buttons, plus page numbers
        #>
        "pagingType"     = $PagingStyle
        "lengthMenu"     = @(
            , @($PagingOptions + (-1))
            , @($PagingOptions + "All")
        )
        "ordering"       = -not $DisableOrdering.IsPresent
        "order"          = @() # this makes sure there's no default ordering upon start (usually it would be 1st column)
        "rowGroup"       = ''
        "info"           = -not $DisableInfo.IsPresent
        "procesing"      = -not $DisableProcessing.IsPresent
        "select"         = -not $DisableSelect.IsPresent
        "searching"      = -not $DisableSearch.IsPresent
        "stateSave"      = -not $DisableStateSave.IsPresent
    }
    if ($ScrollX) {
        $Options.'scrollX' = $true
        # disabling responsive table because it won't work with ScrollX
        $DisableResponsiveTable = $true
    }
    if ($ScrollY) {
        $Options.'scrollY' = "$($ScrollSizeY)px"
    }

    if ($FreezeColumnsLeft -or $FreezeColumnsRight) {
        $Options.fixedColumns = [ordered] @{ }
        if ($FreezeColumnsLeft) {
            $Options.fixedColumns.leftColumns = $FreezeColumnsLeft
        }
        if ($FreezeColumnsRight) {
            $Options.fixedColumns.rightColumns = $FreezeColumnsRight
        }
    }
    if ($FixedHeader -or $FixedFooter) {
        # Using FixedHeader/FixedFooter won't work with ScrollY.
        $Options.fixedHeader = [ordered] @{ }
        if ($FixedHeader) {
            $Options.fixedHeader.header = $FixedHeader.IsPresent
        }
        if ($FixedFooter) {
            $Options.fixedHeader.footer = $FixedFooter.IsPresent
        }
    }
    #}

    # this was due to: https://github.com/DataTables/DataTablesSrc/issues/143
    if (-not $DisableResponsiveTable) {
        $Options["responsive"] = @{ }
        $Options["responsive"]['details'] = @{ }
        if ($ImmediatelyShowHiddenDetails) {
            $Options["responsive"]['details']['display'] = '$.fn.dataTable.Responsive.display.childRowImmediate'
        }
        if ($HideShowButton) {
            $Options["responsive"]['details']['type'] = 'none' # this makes button invisible
        } else {
            $Options["responsive"]['details']['type'] = 'inline' # this adds a button
        }
    } else {
        # HideSHowButton doesn't work
        # ImmediatelyShowHiddenDetails doesn't work
        # Maybe I should communicate this??
        # Better would be with parametersets but don't want to play now
    }


    if ($OrderMulti) {
        $Options.orderMulti = $OrderMulti.IsPresent
    }
    if ($Find -ne '') {
        $Options.search = @{
            search = $Find
        }
    }

    [int] $RowGroupingColumnID = -1
    if ($RowGrouping.Count -gt 0) {
        if ($RowGrouping.Name) {
            $RowGroupingColumnID = ($HeaderNames).ToLower().IndexOf($RowGrouping.Name.ToLower())
        } else {
            $RowGroupingColumnID = $RowGrouping.ColumnID
        }
        if ($RowGroupingColumnID -ne -1) {
            $ColumnsOrder = , @($RowGroupingColumnID, $RowGrouping.Sorting)
            if ($DefaultSortColumn.Count -gt 0 -or $DefaultSortIndex.Count -gt 0) {
                Write-Warning 'New-HTMLTable - Row grouping sorting overwrites default sorting.'
            }
        } else {
            Write-Warning 'New-HTMLTable - Row grouping disabled. Column name/id not found.'
        }
    } else {
        # Sorting
        if ($DefaultSortOrder -eq 'Ascending') {
            $Sort = 'asc'
        } else {
            $Sort = 'desc'
        }
        if ($DefaultSortColumn.Count -gt 0) {
            $ColumnsOrder = foreach ($Column in $DefaultSortColumn) {
                $DefaultSortingNumber = ($HeaderNames).ToLower().IndexOf($Column.ToLower())
                if ($DefaultSortingNumber -ne - 1) {
                    , @($DefaultSortingNumber, $Sort)
                }
            }

        }
        if ($DefaultSortIndex.Count -gt 0 -and $DefaultSortColumn.Count -eq 0) {
            $ColumnsOrder = foreach ($Column in $DefaultSortIndex) {
                if ($Column -ne - 1) {
                    , @($Column, $Sort)
                }
            }
        }
    }
    if ($ColumnsOrder.Count -gt 0) {
        $Options."order" = @($ColumnsOrder)
        # there seems to be a bug in ordering and colReorder plugin
        # Disabling colReorder
        $Options.colReorder = $false
    }

    # Overwriting table size - screen size in percent. With added Section/Panels it shouldn't be more than 90%
    if ($ScreenSizePercent -gt 0) {
        $Options."scrollY" = "$($ScreenSizePercent)vh"
    }
    if ($null -ne $ConditionalFormatting -and $ConditionalFormatting.Count -gt 0) {
        $Options.createdRow = ''
    }

    if ($ResponsivePriorityOrderIndex -or $ResponsivePriorityOrder) {

        $PriorityOrder = 0

        [Array] $PriorityOrderBinding = @(
            foreach ($_ in $ResponsivePriorityOrder) {
                $Index = [array]::indexof($HeaderNames.ToUpper(), $_.ToUpper())
                if ($Index -ne -1) {
                    @{ responsivePriority = 0; targets = $Index }
                }
            }
            foreach ($_ in $ResponsivePriorityOrderIndex) {
                @{ responsivePriority = 0; targets = $_ }
            }
        )
        $Options.columnDefs = @(
            foreach ($_ in $PriorityOrderBinding) {
                $PriorityOrder++
                $_.responsivePriority = $PriorityOrder
                $_
            }
        )
    }

    $Options = $Options | ConvertTo-Json -Depth 6

    # cleans up $Options for ImmediatelyShowHiddenDetails
    # Since it's JavaScript inside we're basically removing double quotes from JSON in favor of no quotes at all
    # Before: "display": "$.fn.dataTable.Responsive.display.childRowImmediate"
    # After: "display": $.fn.dataTable.Responsive.display.childRowImmediate
    $Options = $Options -replace '"(\$\.fn\.dataTable\.Responsive\.display\.childRowImmediate)"', '$1'

    # Process Conditional Formatting. Ugly JS building
    $Options = New-TableConditionalFormatting -Options $Options -ConditionalFormatting $ConditionalFormatting -Header $HeaderNames
    # Process Row Grouping. Ugly JS building
    if ($RowGroupingColumnID -ne -1) {
        $Options = Convert-TableRowGrouping -Options $Options -RowGroupingColumnID $RowGroupingColumnID
        $RowGroupingTop = Add-TableRowGrouping -DataTableName $DataTableID -Top -Settings $RowGrouping
        $RowGroupingBottom = Add-TableRowGrouping -DataTableName $DataTableID -Bottom -Settings $RowGrouping
    }

    [Array] $Tabs = ($Script:HTMLSchema.TabsHeaders | Where-Object { $_.Current -eq $true })
    if ($Tabs.Count -eq 0) {
        # There are no tabs in use, pretend there is only one Active Tab
        $Tab = @{ Active = $true }
    } else {
        # Get First Tab
        $Tab = $Tabs[0]
    }

    # return data
    if (-not $Simplify) {
        $Script:HTMLSchema.Features.Jquery = $true
        $Script:HTMLSchema.Features.DataTables = $true
        $Script:HTMLSchema.Features.DataTablesPDF = $true
        $Script:HTMLSchema.Features.DataTablesExcel = $true
        #$Script:HTMLSchema.Features.DataTablesSearchFade = $true

        if ($ScrollX) {
            $TableAttributes = @{ id = $DataTableID; class = "$($Style -join ' ')"; width = '100%' }
        } else {
            $TableAttributes = @{ id = $DataTableID; class = "$($Style -join ' ')"; width = '100%' }
        }

        # Enable Custom Date fromat sorting
        $SortingFormatDateTime = Add-CustomFormatForDatetimeSorting -DateTimeSortingFormat $DateTimeSortingFormat
        $FilteringOutput = Add-TableFiltering -Filtering $Filtering -FilteringLocation $FilteringLocation -DataTableName $DataTableID
        $FilteringTopCode = $FilteringOutput.FilteringTopCode
        $FilteringBottomCode = $FilteringOutput.FilteringBottomCode
        $LoadSavedState = Add-TableState -DataTableName $DataTableID -Filtering $Filtering -FilteringLocation $FilteringLocation -SavedState (-not $DisableStateSave)

        if ($Tab.Active -eq $true) {
            New-HTMLTag -Tag 'script' {
                @"
                `$(document).ready(function() {
                    $SortingFormatDateTime
                    $RowGroupingTop
                    $LoadSavedState
                    $FilteringTopCode
                    //  Table code
                    var table = `$('#$DataTableID').DataTable(
                        $($Options)
                    );
                    $FilteringBottomCode
                    $RowGroupingBottom
                });
"@
            }

        } else {
            [string] $TabName = $Tab.Id
            New-HTMLTag -Tag 'script' {
                @"
                    `$(document).ready(function() {
                        $SortingFormatDateTime
                        $RowGroupingTop
                        `$('.tabs').on('click', 'a', function (event) {
                            if (`$(event.currentTarget).attr("data-id") == "$TabName" && !$.fn.dataTable.isDataTable("#$DataTableID")) {
                                $LoadSavedState
                                $FilteringTopCode
                                //  Table code
                                var table = `$('#$DataTableID').DataTable(
                                    $($Options)
                                );
                                $FilteringBottomCode
                            };
                        });
                        $RowGroupingBottom
                    });
"@
            }
        }
    } else {
        $TableAttributes = @{ class = 'simplify' }
        $Script:HTMLSchema.Features.DataTablesSimplify = $true
    }

    if ($InvokeHTMLTags) {
        # By default HTML tags are displayed, in this case we're converting tags into real tags
        $Table = $Table -replace '&lt;', '<' -replace '&gt;', '>' -replace '&amp;nbsp;', ' ' -replace '&quot;', '"' -replace '&#39;', "'"
    }
    if (-not $DisableNewLine) {
        # Finds new lines and adds HTML TAG BR
        #$Table = $Table -replace '(?m)\s+$', "`r`n<BR>"
        $Table = $Table -replace '(?m)\s+$', "<BR>"
    }

    if ($OtherHTML) {
        $BeforeTableCode = Invoke-Command -ScriptBlock $OtherHTML
    } else {
        $BeforeTableCode = ''
    }

    if ($PreContent) {
        $BeforeTable = Invoke-Command -ScriptBlock $PreContent
    } else {
        $BeforeTable = ''
    }
    if ($PostContent) {
        $AfterTable = Invoke-Command -ScriptBlock $PostContent
    } else {
        $AfterTable = ''
    }

    if ($RowGrouping.Attributes.Count -gt 0) {
        $RowGroupingCSS = ConvertTo-CSS -ID $DataTableID -ClassName 'tr.dtrg-group td' -Attributes $RowGrouping.Attributes -Group
    } else {
        $RowGroupingCSS = ''
    }

    New-HTMLTag -Tag 'div' -Attributes @{ class = 'flexElement overflowHidden' } -Value {
        $RowGroupingCSS
        $BeforeTableCode
        $BeforeTable
        # Build HTML TABLE
        New-HTMLTag -Tag 'table' -Attributes $TableAttributes {
            New-HTMLTag -Tag 'thead' {
                if ($AddedHeader) {
                    $AddedHeader
                } else {
                    $Header
                }
            }
            New-HTMLTag -Tag 'tbody' {
                $Table
            }
            if (-not $HideFooter) {
                New-HTMLTag -Tag 'tfoot' {
                    $Header
                }
            }
        }
        $AfterTable
    }
}

function New-HTMLTabOptions {
    [alias('TabOptions')]
    [CmdletBinding()]
    param(
        [switch] $SlimTabs,
        [string] $SelectorColor,
        [string] $SelectorColorTarget,
        [switch] $Transition,
        [switch] $LinearGradient,
        [ValidateSet('0px', '10px', '15px', '25px')][string] $BorderRadius = '0px',
        [string] $BorderBackgroundColor

    )
    if (-not $Script:HTMLSchema) {
        Write-Warning 'New-HTMLTabOptions - Creation of HTML aborted. Most likely New-HTML is missing.'
        Exit
    }
    #$Script:HTMLSchema.TabOptions = @{ }
    $Script:HTMLSchema.TabOptions.SlimTabs = $SlimTabs.IsPresent
    if ($SelectorColor) {
        # $Script:HTMLSchema.TabOptions.SelectorColor = ConvertFrom-Color -Color $SelectorColor
        $Script:Configuration.Features.Tabbis.CustomActionsReplace.ColorSelector = ConvertFrom-Color -Color $SelectorColor
        $Script:Configuration.Features.TabbisGradient.CustomActionsReplace.ColorSelector = ConvertFrom-Color -Color $SelectorColor
        # $Script:Configuration.Features.TabsTransition.CustomActionsReplace.ColorSelector = ConvertFrom-Color -Color $SelectorColor
    }
    if ($SelectorColorTarget) {
        $Script:Configuration.Features.Tabbis.CustomActionsReplace.ColorTarget = ConvertFrom-Color -Color $SelectorColorTarget
        $Script:Configuration.Features.TabbisGradient.CustomActionsReplace.ColorTarget = ConvertFrom-Color -Color $SelectorColorTarget
    }
    $Script:HTMLSchema.Features.TabbisGradient = $LinearGradient.IsPresent
    $Script:HTMLSchema.Features.TabbisTransition = $Transition.IsPresent
    
    $Script:BorderStyle = @{
        'border-radius'    = "$BorderRadius";
        'background-color' = ""
    }

    if ($BorderBackgroundColor) {
        $Script:BorderStyle.'background-color' = ConvertFrom-Color -Color $BorderBackgroundColor
    }
}

Register-ArgumentCompleter -CommandName New-HTMLTabOptions -ParameterName SelectorColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLTabOptions -ParameterName SelectorColorTarget -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLTabOptions -ParameterName BorderBackgroundColor -ScriptBlock { $Script:RGBColors.Keys }

function New-HTMLTag {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)][alias('Content')][ScriptBlock] $Value,
        [Parameter(Mandatory = $true, Position = 1)][string] $Tag,
        [System.Collections.IDictionary] $Attributes,
        [switch] $SelfClosing,
        [switch] $NewLine
    )
    $HTMLTag = [Ordered] @{
        Tag         = $Tag
        Attributes  = $Attributes
        Value       = if ($null -eq $Value) { '' } else { Invoke-Command -ScriptBlock $Value }
        SelfClosing = $SelfClosing
    }
    $HTML = Set-Tag -HtmlObject $HTMLTag -NewLine:$NewLine
    return $HTML
}
function New-HTMLText {
    [alias('HTMLText', 'Text')]
    [CmdletBinding()]
    param(
        [string[]] $Text,
        [string[]] $Color = @(),
        [string[]] $BackGroundColor = @(),
        [int[]] $FontSize = @(),
        [ValidateSet('normal', 'bold', 'bolder', 'lighter', '100', '200', '300', '400', '500', '600', '700', '800', '900')][string[]] $FontWeight = @(),
        [ValidateSet('normal', 'italic', 'oblique')][string[]] $FontStyle = @(),
        [ValidateSet('normal', 'small-caps')][string[]] $FontVariant = @(),
        [string[]] $FontFamily = @(),
        [ValidateSet('left', 'center', 'right', 'justify')][string[]] $Alignment = @(),
        [ValidateSet('none', 'line-through', 'overline', 'underline')][string[]] $TextDecoration = @(),
        [ValidateSet('uppercase', 'lowercase', 'capitalize')][string[]] $TextTransform = @(),
        [ValidateSet('rtl')][string[]] $Direction = @(),
        [switch] $LineBreak,
        [switch] $SkipParagraph #,
        #[bool[]] $NewLine = @()
    )
    #Write-Verbose 'New-HTMLText - Processing...'
    $DefaultColor = $Color[0]
    $DefaultFontSize = $FontSize[0]
    $DefaultFontWeight = if ($null -eq $FontWeight[0] ) { '' } else { $FontWeight[0] }
    $DefaultBackGroundColor = $BackGroundColor[0]
    $DefaultFontFamily = if ($null -eq $FontFamily[0] ) { '' } else { $FontFamily[0] }
    $DefaultFontStyle = if ($null -eq $FontStyle[0] ) { '' } else { $FontStyle[0] }
    $DefaultTextDecoration = if ($null -eq $TextDecoration[0]) { '' } else { $TextDecoration[0] }
    $DefaultTextTransform = if ($null -eq $TextTransform[0]) { '' } else { $TextTransform[0] }
    $DefaultFontVariant = if ($null -eq $FontVariant[0]) { '' } else { $FontVariant }
    $DefaultDirection = if ($null -eq $Direction[0]) { '' } else { $Direction[0] }
    $DefaultAlignment = if ($null -eq $Alignment[0]) { '' } else { $Alignment[0] }
    # $DefaultNewLine = if ($null -eq $NewLine[0]) { $false } else { $NewLine[0] }

    $Output = for ($i = 0; $i -lt $Text.Count; $i++) {
        if ($null -eq $FontWeight[$i]) {
            $ParamFontWeight = $DefaultFontWeight
        } else {
            $ParamFontWeight = $FontWeight[$i]
        }
        if ($null -eq $FontSize[$i]) {
            $ParamFontSize = $DefaultFontSize
        } else {
            $ParamFontSize = $FontSize[$i]
        }
        if ($null -eq $Color[$i]) {
            $ParamColor = $DefaultColor
        } else {
            $ParamColor = $Color[$i]
        }
        if ($null -eq $BackGroundColor[$i]) {
            $ParamBackGroundColor = $DefaultBackGroundColor
        } else {
            $ParamBackGroundColor = $BackGroundColor[$i]
        }
        if ($null -eq $FontFamily[$i]) {
            $ParamFontFamily = $DefaultFontFamily
        } else {
            $ParamFontFamily = $FontFamily[$i]
        }
        if ($null -eq $FontStyle[$i]) {
            $ParamFontStyle = $DefaultFontStyle
        } else {
            $ParamFontStyle = $FontStyle[$i]
        }

        if ($null -eq $TextDecoration[$i]) {
            $ParamTextDecoration = $DefaultTextDecoration
        } else {
            $ParamTextDecoration = $TextDecoration[$i]
        }

        if ($null -eq $TextTransform[$i]) {
            $ParamTextTransform = $DefaultTextTransform
        } else {
            $ParamTextTransform = $TextTransform[$i]
        }

        if ($null -eq $FontVariant[$i]) {
            $ParamFontVariant = $DefaultFontVariant
        } else {
            $ParamFontVariant = $FontVariant[$i]
        }
        if ($null -eq $Direction[$i]) {
            $ParamDirection = $DefaultDirection
        } else {
            $ParamDirection = $Direction[$i]
        }
        if ($null -eq $Alignment[$i]) {
            $ParamAlignment = $DefaultAlignment
        } else {
            $ParamAlignment = $Alignment[$i]
        }

        $newSpanTextSplat = @{ }
        $newSpanTextSplat.Color = $ParamColor
        $newSpanTextSplat.BackGroundColor = $ParamBackGroundColor

        $newSpanTextSplat.FontSize = $ParamFontSize
        if ($ParamFontWeight -ne '') {
            $newSpanTextSplat.FontWeight = $ParamFontWeight
        }
        $newSpanTextSplat.FontFamily = $ParamFontFamily
        if ($ParamFontStyle -ne '') {
            $newSpanTextSplat.FontStyle = $ParamFontStyle
        }
        if ($ParamFontVariant -ne '') {
            $newSpanTextSplat.FontVariant = $ParamFontVariant
        }
        if ($ParamTextDecoration -ne '') {
            $newSpanTextSplat.TextDecoration = $ParamTextDecoration
        }
        if ($ParamTextTransform -ne '') {
            $newSpanTextSplat.TextTransform = $ParamTextTransform
        }
        if ($ParamDirection -ne '') {
            $newSpanTextSplat.Direction = $ParamDirection
        }
        if ($ParamAlignment -ne '') {
            $newSpanTextSplat.Alignment = $ParamAlignment
        }

        $newSpanTextSplat.LineBreak = $LineBreak
        New-HTMLSpanStyle @newSpanTextSplat {
            if ($Text[$i] -match "\[([^\[]*)\)") {
                # Covers markdown LINK  "[somestring](https://evotec.xyz)"
                $RegexBrackets1 = [regex] "\[([^\[]*)\]" # catch 'sometstring'
                $RegexBrackets2 = [regex] "\(([^\[]*)\)" # catch link
                $RegexBrackets3 = [regex] "\[([^\[]*)\)" # catch both somestring and link
                $Text1 = $RegexBrackets1.match($Text[$i]).Groups[1].value
                $Text2 = $RegexBrackets2.match($Text[$i]).Groups[1].value
                $Text3 = $RegexBrackets3.match($Text[$i]).Groups[0].value
                if ($Text1 -ne '' -and $Text2 -ne '') {
                    $Link = New-HTMLAnchor -HrefLink $Text2 -Text $Text1
                    $Text[$i].Replace($Text3, $Link)
                }
            } else {
                # Default
                $Text[$i]
                # if ($NewLine[$i]) {
                #    '<br>'
                #}
            }
        }
    }

    if ($SkipParagraph) {
        $Output -join ''
    } else {
        New-HTMLTag -Tag 'div' {
            $Output
        }
    }
    if ($LineBreak) {
        New-HTMLTag -Tag 'br' -SelfClosing
    }
}

Register-ArgumentCompleter -CommandName New-HTMLText -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLText -ParameterName BackgroundColor -ScriptBlock { $Script:RGBColors.Keys }
function New-HTMLTimeline {
    param(
        [Parameter(Mandatory = $false, Position = 0)][alias('TimeLineItems')][ScriptBlock] $Content
    )
    $Script:HTMLSchema.Features.TimeLine = $true
    New-HTMLTag -Tag 'div' -Attributes @{ class = 'timelineSimpleContainer' } {
        if ($null -eq $Value) { '' } else { Invoke-Command -ScriptBlock $Content }
    }
}
function New-HTMLTimelineItem {
    [CmdletBinding()]
    param(
        [DateTime] $Date = (Get-Date),
        [string] $HeadingText,
        [string] $Text,
        [string] $Color
    )
    $Attributes = @{
        class     = 'timelineSimple-item'
        "date-is" = $Date
    }

    if ($null -ne $Color) {
        $RGBcolor = ConvertFrom-Color -Color $Color
        $Style = "color: $RGBcolor;"
    } else {
        $Style = ''
    }
    # $Script:HTMLSchema.Features.TimeLine = $true
    New-HTMLTag -Tag 'div' -Attributes $Attributes -Value {
        New-HTMLTag -Tag 'h1' -Attributes @{ class = 'timelineSimple'; style = $style } {
            $HeadingText
        }
        New-HTMLTag -Tag 'p' -Attributes @{ class = 'timelineSimple' } {
            $Text -Replace [Environment]::NewLine, '<br>' -replace '\n', '<br>'
        }
    }
}

Register-ArgumentCompleter -CommandName New-HTMLTimelineItem -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
function New-HTMLToast {
    [CmdletBinding()]
    param(
        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")][string] $TextHeader,

        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")][string] $TextHeaderColor,

        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")][string] $Text,

        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")][string] $TextColor,

        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")][int] $IconSize = 30,

        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")][string] $IconColor = "Blue",

        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")][string] $BarColorLeft = "Blue",

        [parameter(ParameterSetName = "FontAwesomeBrands")]
        [parameter(ParameterSetName = "FontAwesomeRegular")]
        [parameter(ParameterSetName = "FontAwesomeSolid")][string] $BarColorRight,

        # ICON BRANDS
        [ArgumentCompleter(
            {
                param($CommandName, $ParameterName, $WordToComplete, $CommandAst, $FakeBoundParameters)
                ($Global:HTMLIcons.FontAwesomeBrands.Keys)
            }
        )]
        [ValidateScript(
            {
                $_ -in (($Global:HTMLIcons.FontAwesomeBrands.Keys))
            }
        )]
        [parameter(ParameterSetName = "FontAwesomeBrands")][string] $IconBrands,

        # ICON REGULAR
        [ArgumentCompleter(
            {
                param($CommandName, $ParameterName, $WordToComplete, $CommandAst, $FakeBoundParameters)
                ($Global:HTMLIcons.FontAwesomeRegular.Keys)
            }
        )]
        [ValidateScript(
            {
                $_ -in (($Global:HTMLIcons.FontAwesomeRegular.Keys))
            }
        )]
        [parameter(ParameterSetName = "FontAwesomeRegular")][string] $IconRegular,

        # ICON SOLID
        [ArgumentCompleter(
            {
                param($CommandName, $ParameterName, $WordToComplete, $CommandAst, $FakeBoundParameters)
                ($Global:HTMLIcons.FontAwesomeSolid.Keys)
            }
        )]
        [ValidateScript(
            {
                $_ -in (($Global:HTMLIcons.FontAwesomeSolid.Keys))
            }
        )]
        [parameter(ParameterSetName = "FontAwesomeSolid")][string] $IconSolid
    )

    [string] $Icon = ''
    if ($IconBrands) {
        $Icon = "fab fa-$IconBrands" # fa-$($FontSize)x"
    } elseif ($IconRegular) {
        $Icon = "far fa-$IconRegular" # fa-$($FontSize)x"
    } elseif ($IconSolid) {
        $Icon = "fas fa-$IconSolid" # fa-$($FontSize)x"
    }

    $Script:HTMLSchema.Features.Toasts = $true

    [string] $DivClass = "toast"

    $StyleText = @{ }
    if ($TextColor) {
        $StyleText.'color' = ConvertFrom-Color -Color $TextColor
    }

    $StyleTextHeader = @{ }
    if ($TextHeaderColor) {
        $StyleTextHeader.'color' = ConvertFrom-Color -Color $TextHeaderColor
    }

    $StyleIcon = @{ }
    if ($IconSize -ne 0) {
        $StyleIcon.'font-size' = "$($IconSize)px"
    }

    if ($IconColor) {
        $StyleIcon.'color' = ConvertFrom-Color -Color $IconColor
    }

    $StyleBarLeft = @{ }
    if ($BarColorLeft) {
        $StyleBarLeft.'background-color' = ConvertFrom-Color -Color $BarColorLeft
    }

    $StyleBarRight = @{ }
    if ($BarColorRight) {
        $StyleBarRight.'background-color' = ConvertFrom-Color -Color $BarColorRight
    }

    New-HTMLTag -Tag 'div' -Attributes @{ class = $DivClass } {
        New-HTMLTag -Tag 'div' -Attributes @{ class = 'toastBorderLeft'; style = $StyleBarLeft }
        New-HTMLTag -Tag 'div' -Attributes @{ class = "toastIcon $Icon"; style = $StyleIcon }
        New-HTMLTag -Tag 'div' -Attributes @{ class = 'toastContent' } {
            New-HTMLTag -Tag 'p' -Attributes @{ class = 'toastTextHeader'; style = $StyleTextHeader } {
                $TextHeader
            }
            New-HTMLTag -Tag 'p' -Attributes @{ class = 'toastText'; style = $StyleText } {
                $Text
            }
        }
        New-HTMLTag -Tag 'div' -Attributes @{ class = 'toastBorderRight'; style = $StyleBarRight }
    }
}

Register-ArgumentCompleter -CommandName New-HTMLToast -ParameterName TextHeaderColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLToast -ParameterName TextColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLToast -ParameterName IconColor -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLToast -ParameterName BarColorLeft -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-HTMLToast -ParameterName BarColorRight -ScriptBlock { $Script:RGBColors.Keys }
function New-HTMLTree {
    [CmdletBinding()]
    param(
        [scriptblock] $Nodes,
        [ValidateSet('none', 'checkbox', 'radio')][string] $Checkbox = 'none',
        [ValidateSet('none', '1', '2', '3')] $SelectMode = '2'
    )
    $Script:HTMLSchema.Features.Jquery = $true
    $Script:HTMLSchema.Features.FancyTree = $true

    [string] $ID = "FancyTree" + (Get-RandomStringName -Size 8)

    $FancyTree = @{

    }

    $FancyTree['extensions'] = @("edit", "filter")

    if ($Checkbox -eq 'none') {
        #$FancyTree['checkbox'] = $Checkbox # true/false/radio
    } elseif ($Checkbox -eq 'radio') {
        $FancyTree['checkbox'] = 'radio'
    } else {
        $FancyTree['checkbox'] = $true
    }
    <#
    Fancytree supports three modes:

    selectMode: 1: single selection, Only one node is selected at any time.
    selectMode: 2: multiple selection (default), Every node may be selected independently.
    selectMode: 3: hierarchical selection, (De)selecting a node will propagate to all descendants. Mixed states will be displayed as partially selected using a tri-state checkbox.
    #>
    if ($SelectMode -ne 'none') {
        $FancyTree['selectMode'] = $SelectMode # 3, // 1, 2, 3
    }

    [Array] $Source = & $Nodes
    if ($Source.Count -gt 0) {
        $FancyTree['source'] = $Source
    }
    Remove-EmptyValues -Hashtable $FancyTree -Rerun 1 -Recursive

    # Build HTML
    $Div = New-HTMLTag -Tag 'div' -Attributes @{ id = $ID; }

    $Script = New-HTMLTag -Tag 'script' -Value {
        $DivID = -join ('#', $ID)
        # Convert Dictionary to JSON and return chart within SCRIPT tag
        # Make sure to return with additional empty string
        $FancyTreeJSON = $FancyTree | ConvertTo-Json -Depth 100 | ForEach-Object { [System.Text.RegularExpressions.Regex]::Unescape($_) }

        '$(function(){  // on page load'
        "`$(`"$DivID`").fancytree("
        $FancyTreeJSON
        ');'
        '});'
    } -NewLine

    # Return Data
    $Div
    $Script
}
function New-TableButtonCopy {
    [alias('TableButtonCopy', 'EmailTableButtonCopy', 'New-HTMLTableButtonCopy')]
    [CmdletBinding()]
    param()

    [PSCustomObject] @{
        Type   = 'TableButtonCopy'
        Output = @{
            extend = 'copyHtml5'
        }
    }
}
function New-TableButtonCSV {
    [alias('TableButtonCSV', 'EmailTableButtonCSV', 'New-HTMLTableButtonCSV')]
    [CmdletBinding()]
    param()
    [PSCustomObject] @{
        Type   = 'TableButtonCSV'
        Output = @{
            extend = 'csvHtml5'
        }
    }
}
function New-TableButtonExcel {
    [alias('TableButtonExcel', 'EmailTableButtonExcel', 'New-HTMLTableButtonExcel')]
    [CmdletBinding()]
    param()
    [PSCustomObject] @{
        Type   = 'TableButtonExcel'
        Output = @{
            extend = 'excelHtml5'
        }
    }
}
function New-TableButtonPageLength {
    [alias('TableButtonPageLength', 'EmailTableButtonPageLength', 'New-HTMLTableButtonPageLength')]
    [CmdletBinding()]
    param()
    [PSCustomObject] @{
        Type   = 'TableButtonPageLength'
        Output = @{
            extend = 'pageLength'
        }
    }
}
function New-TableButtonPDF {
    <#
    .SYNOPSIS
    Allows more control when adding buttons to Table

    .DESCRIPTION
    Allows more control when adding buttons to Table. Works only within Table or New-HTMLTable scriptblock.

    .PARAMETER Title
    Document title (appears above the table in the generated PDF). The special character * is automatically replaced with the value read from the host document's title tag.

    .PARAMETER DisplayName
    The button's display text. The text can be configured using this option

    .PARAMETER MessageBottom
    Message to be shown at the bottom of the table, or the caption tag if displayed at the bottom of the table.

    .PARAMETER MessageTop
    Message to be shown at the top of the table, or the caption tag if displayed at the top of the table.

    .PARAMETER FileName
    File name to give the created file (plus the extension defined by the extension option). The special character * is automatically replaced with the value read from the host document's title tag.

    .PARAMETER Extension
    The extension to give the created file name. (default .pdf)

    .PARAMETER PageSize
    Paper size for the created PDF. This can be A3, A4, A5, LEGAL, LETTER or TABLOID. Other options are available.

    .PARAMETER Orientation
    Paper orientation for the created PDF. This can be portrait or landscape

    .PARAMETER Header
    Indicate if the table header should be included in the exported data or not.

    .PARAMETER Footer
    Indicate if the table footer should be included in the exported data or not.

    .EXAMPLE
    Dashboard -Name 'Dashimo Test' -FilePath $PSScriptRoot\DashboardEasy05.html -Show {
        Section -Name 'Test' -Collapsable {
            Container {
                Panel {
                    Table -DataTable $Process {
                        TableButtonPDF
                        TableButtonCopy
                        TableButtonExcel
                    } -Buttons @() -DisableSearch -DisablePaging -HideFooter
                }
                Panel {
                    Table -DataTable $Process -Buttons @() -DisableSearch -DisablePaging -HideFooter
                }
                Panel {
                    Table -DataTable $Process {
                        TableButtonPDF -PageSize A10 -Orientation landscape
                        TableButtonCopy
                        TableButtonExcel
                    } -Buttons @() -DisableSearch -DisablePaging -HideFooter
                }
            }
        }
    }

    .NOTES
    Options are based on this URL: https://datatables.net/reference/button/pdfHtml5

    #>

    [alias('TableButtonPDF', 'EmailTableButtonPDF', 'New-HTMLTableButtonPDF')]
    [CmdletBinding()]
    param(
        [string] $Title,
        [string] $DisplayName,
        [string] $MessageBottom,
        [string] $MessageTop,
        [string] $FileName,
        [string] $Extension,
        [string][ValidateSet('4A0', '2A0', 'A0', 'A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9', 'A10',
            'B0', 'B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'B10',
            'C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9', 'C10',
            'RA0', 'RA1', 'RA2', 'RA3', 'RA4',
            'SRA0', 'SRA1', 'SRA2', 'SRA3', 'SRA4',
            'EXECUTIVE', 'FOLIO', 'LEGAL', 'LETTER', 'TABLOID')] $PageSize = 'A3',
        [string][ValidateSet('portrait', 'landscape')] $Orientation = 'landscape',
        [switch] $Header,
        [switch] $Footer
    )
    $Button = @{ }
    $Button.extend = 'pdfHtml5'
    $Button.pageSize = $PageSize
    $Button.orientation = $Orientation
    if ($MessageBottom) {
        $Button.messageBottom = $MessageBottom
    }
    if ($MessageTop) {
        $Button.messageTop = $MessageTop
    }
    if ($DisplayName) {
        $Button.text = $DisplayName
    }
    if ($Title) {
        $Button.title = $Title
    }
    if ($FileName) {
        $Button.filename = $FileName
    }
    if ($Extension) {
        $Button.extension = $Extension
    }
    if ($Header) {
        $Button.header = $Header.IsPresent
    }
    if ($Footer) {
        $Button.footer = $Footer.IsPresent
    }

    [PSCustomObject] @{
        Type   = 'TableButtonPDF'
        Output = $Button
    }
}







function New-TableButtonPrint {
    [alias('TableButtonPrint', 'EmailTableButtonPrint', 'New-HTMLTableButtonPrint')]
    [CmdletBinding()]
    param()
    $Button = @{
        extend = 'print'
    }
    [PSCustomObject] @{
        Type   = 'TableButtonPrint'
        Output = $Button
    }
}
function New-TableCondition {
    [alias('EmailTableCondition', 'TableConditionalFormatting', 'New-HTMLTableCondition', 'TableCondition')]
    [CmdletBinding()]
    param(
        [alias('ColumnName')][string] $Name,
        [alias('Type')][ValidateSet('number', 'string')][string] $ComparisonType,
        [ValidateSet('lt', 'le', 'eq', 'ge', 'gt', 'ne', 'contains', 'like')][string] $Operator,
        [Object] $Value,
        [switch] $Row,
        [string]$Color,
        [string]$BackgroundColor,
        [int] $FontSize,
        [ValidateSet('normal', 'bold', 'bolder', 'lighter', '100', '200', '300', '400', '500', '600', '700', '800', '900')][string] $FontWeight,
        [ValidateSet('normal', 'italic', 'oblique')][string] $FontStyle,
        [ValidateSet('normal', 'small-caps')][string] $FontVariant,
        [string] $FontFamily,
        [ValidateSet('left', 'center', 'right', 'justify')][string] $Alignment,
        [ValidateSet('none', 'line-through', 'overline', 'underline')][string] $TextDecoration,
        [ValidateSet('uppercase', 'lowercase', 'capitalize')][string] $TextTransform,
        [ValidateSet('rtl')][string] $Direction,
        [switch] $Inline

    )
    $Style = @{
        Color           = $Color
        BackGroundColor = $BackGroundColor
        FontSize        = $FontSize
        FontWeight      = $FontWeight
        FontStyle       = $FontStyle
        FontVariant     = $FontVariant
        FontFamily      = $FontFamily
        Alignment       = $Alignment
        TextDecoration  = $TextDecoration
        TextTransform   = $TextTransform
        Direction       = $Direction
    }
    Remove-EmptyValues -Hashtable $Style

    $TableCondition = [PSCustomObject] @{
        Row             = $Row
        Type            = if (-not $ComparisonType) { 'string' } else { $ComparisonType }
        Name            = $Name
        Operator        = if (-not $Operator) { 'eq' } else { $Operator }
        Value           = $Value
        Color           = $Color
        BackgroundColor = $BackgroundColor
        Style           = ConvertTo-HTMLStyle @Style
    }
    [PSCustomObject] @{
        Type   = if ($Inline) { 'TableConditionInline' } else { 'TableCondition' }
        Output = $TableCondition
    }
}

Register-ArgumentCompleter -CommandName New-TableCondition -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-TableCondition -ParameterName BackgroundColor -ScriptBlock { $Script:RGBColors.Keys }
function New-TableContent {
    [alias('TableContent', 'EmailTableContent', 'New-HTMLTableContent')]
    [CmdletBinding()]
    param(
        [alias('ColumnNames', 'Names', 'Name')][string[]] $ColumnName,
        [int[]] $ColumnIndex,
        [int[]] $RowIndex,
        [string[]] $Text,
        [string] $Color,
        [string] $BackGroundColor,
        [int] $FontSize,
        [ValidateSet('normal', 'bold', 'bolder', 'lighter', '100', '200', '300', '400', '500', '600', '700', '800', '900')][string] $FontWeight,
        [ValidateSet('normal', 'italic', 'oblique')][string] $FontStyle,
        [ValidateSet('normal', 'small-caps')][string] $FontVariant,
        [string] $FontFamily,
        [ValidateSet('left', 'center', 'right', 'justify')][string] $Alignment,
        [ValidateSet('none', 'line-through', 'overline', 'underline')][string] $TextDecoration,
        [ValidateSet('uppercase', 'lowercase', 'capitalize')][string] $TextTransform,
        [ValidateSet('rtl')][string] $Direction
    )

    $Style = @{
        Color           = $Color
        BackGroundColor = $BackGroundColor
        FontSize        = $FontSize
        FontWeight      = $FontWeight
        FontStyle       = $FontStyle
        FontVariant     = $FontVariant
        FontFamily      = $FontFamily
        Alignment       = $Alignment
        TextDecoration  = $TextDecoration
        TextTransform   = $TextTransform
        Direction       = $Direction
    }
    Remove-EmptyValues -Hashtable $Style

    [PSCustomObject]@{
        Type   = 'TableContentStyle'
        Output = @{
            Name        = $ColumnName
            Text        = $Text
            RowIndex    = $RowIndex | Sort-Object
            ColumnIndex = $ColumnIndex | Sort-Object
            Style       = ConvertTo-HTMLStyle @Style
            Used        = $false
        }
    }
}

Register-ArgumentCompleter -CommandName New-TableContent -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-TableContent -ParameterName BackgroundColor -ScriptBlock { $Script:RGBColors.Keys }
function New-TableHeader {
    [alias('TableHeader', 'EmailTableHeader', 'New-HTMLTableHeader')]
    [CmdletBinding()]
    param(
        [string[]] $Names,
        [string] $Title,
        [string] $Color,
        [string] $BackGroundColor,
        [int] $FontSize,
        [ValidateSet('normal', 'bold', 'bolder', 'lighter', '100', '200', '300', '400', '500', '600', '700', '800', '900')][string] $FontWeight,
        [ValidateSet('normal', 'italic', 'oblique')][string] $FontStyle,
        [ValidateSet('normal', 'small-caps')][string] $FontVariant,
        [string] $FontFamily,
        [ValidateSet('left', 'center', 'right', 'justify')][string] $Alignment,
        [ValidateSet('none', 'line-through', 'overline', 'underline')][string] $TextDecoration,
        [ValidateSet('uppercase', 'lowercase', 'capitalize')][string] $TextTransform,
        [ValidateSet('rtl')][string] $Direction,
        [switch] $AddRow,
        [int] $ColumnCount,
        [ValidateSet(
            'all',
            'none',
            'never',
            'desktop',
            'not-desktop',
            'tablet-l',
            'tablet-p',
            'mobile-l',
            'mobile-p',
            'min-desktop',
            'max-desktop',
            'tablet',
            'not-tablet',
            'min-tablet',
            'max-tablet',
            'not-tablet-l',
            'min-tablet-l',
            'max-tablet-l',
            'not-tablet-p',
            'min-tablet-p',
            'max-tablet-p',
            'mobile',
            'not-mobile',
            'min-mobile',
            'max-mobile',
            'not-mobile-l',
            'min-mobile-l',
            'max-mobile-l',
            'not-mobile-p',
            'min-mobile-p',
            'max-mobile-p'
        )][string] $ResponsiveOperations

    )
    if ($AddRow) {
        Write-Warning "New-HTMLTableHeader - Using AddRow switch is deprecated. It's not nessecary anymore. Just use Title alone. It will be removed later on."
    }

    $Style = @{
        Color           = $Color
        BackGroundColor = $BackGroundColor
        FontSize        = $FontSize
        FontWeight      = $FontWeight
        FontStyle       = $FontStyle
        FontVariant     = $FontVariant
        FontFamily      = $FontFamily
        Alignment       = $Alignment
        TextDecoration  = $TextDecoration
        TextTransform   = $TextTransform
        Direction       = $Direction
    }
    Remove-EmptyValues -Hashtable $Style

    if (($AddRow -and $Title) -or ($Title -and -not $Names)) {
        $Type = 'TableHeaderFullRow'
    } elseif ((-not $AddRow -and $Title) -or ($Title -and $Names)) {
        $Type = 'TableHeaderMerge'
    } elseif ($Names -and $ResponsiveOperations) {
        $Type = 'TableHeaderResponsiveOperations'
    } elseif ($ResponsiveOperations) {
        Write-Warning 'New-HTMLTableHeader - ResponsiveOperations require Names (ColumnNames) to apply operation to.'
        return
    } else {
        $Type = 'TableHeaderStyle'
    }

    [PSCustomObject]@{
        Type   = $Type
        Output = @{
            Names                = $Names
            ResponsiveOperations = $ResponsiveOperations
            Title                = $Title
            Style                = ConvertTo-HTMLStyle @Style
            ColumnCount          = $ColumnCount
        }
    }
}

Register-ArgumentCompleter -CommandName New-TableHeader -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-TableHeader -ParameterName BackGroundColor -ScriptBlock { $Script:RGBColors.Keys }
function New-TableReplace {
    [alias('TableReplace', 'EmailTableReplace', 'New-HTMLTableReplace')]
    [CmdletBinding()]
    param(
        [string] $FieldName,
        [string[]] $Replacements

    )
    [PSCustomObject]@{
        Type   = 'TableReplaceCompare'
        Output = @{
            $FieldName = $Replacements
        }
    }
}
function New-TableRowGrouping {
    [alias('TableRowGrouping', 'EmailTableRowGrouping', 'New-HTMLTableRowGrouping')]
    [CmdletBinding()]
    param(
        [alias('ColumnName')][string] $Name,
        [int] $ColumnID = -1,
        [ValidateSet('Ascending', 'Descending')][string] $SortOrder = 'Ascending',
        [string] $Color,
        [string] $BackgroundColor
    )

    $Object = [PSCustomObject] @{
        Type   = 'TableRowGrouping'
        Output = [ordered] @{
            Name       = $Name
            ColumnID   = $ColumnID
            Sorting    = if ('Ascending') { 'asc' } else { 'desc' }
            Attributes = @{
                'color'            = ConvertFrom-Color -Color $Color
                'background-color' = ConvertFrom-Color -Color $BackgroundColor
            }
        }
    }
    Remove-EmptyValues -Hashtable $Object.Output
    $Object
}

Register-ArgumentCompleter -CommandName New-TableRowGrouping -ParameterName Color -ScriptBlock { $Script:RGBColors.Keys }
Register-ArgumentCompleter -CommandName New-TableRowGrouping -ParameterName BackgroundColor -ScriptBlock { $Script:RGBColors.Keys }
function New-TreeNode {
    [CmdletBinding()]
    param(
        [scriptblock] $Children,
        [string] $Title,
        [string] $Id,
        [switch] $Folder
    )

    if ($Children) {
        [Array] $SourceChildren = & $Children
    }

    $Node = [ordered] @{
        title  = $Title
        key    = $Id
        folder = $Folder.IsPresent
    }
    if ($SourceChildren.Count) {
        $Node['children'] = $SourceChildren
    }
    $Node
}
function Out-HtmlView {
    <#
    .SYNOPSIS
    Small function that allows to send output to HTML

    .DESCRIPTION
    Small function that allows to send output to HTML. When displaying in HTML it allows data to output to EXCEL, CSV and PDF. It allows sorting, searching and so on.

    .PARAMETER Table
    Data you want to display

    .PARAMETER Title
    Title of HTML Window

    .PARAMETER DefaultSortColumn
    Sort by Column Name

    .PARAMETER DefaultSortIndex
    Sort by Column Index

    .EXAMPLE
    Get-Process | Select-Object -First 5 | Out-HtmlView

    .NOTES
    General notes
    #>
    [alias('Out-GridHtml', 'ohv')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)][ScriptBlock] $HTML,
        [Parameter(Mandatory = $false, Position = 1)][ScriptBlock] $PreContent,
        [Parameter(Mandatory = $false, Position = 2)][ScriptBlock] $PostContent,
        [alias('ArrayOfObjects', 'Object', 'DataTable')][Parameter(ValueFromPipeline = $true, Mandatory = $true)] $Table,
        [string] $FilePath,
        [string] $Title = 'Out-HTMLView',
        [switch] $PassThru,
        [string[]][ValidateSet('copyHtml5', 'excelHtml5', 'csvHtml5', 'pdfHtml5')] $Buttons = @('copyHtml5', 'excelHtml5', 'csvHtml5', 'pdfHtml5', 'pageLength'),
        [string[]][ValidateSet('numbers', 'simple', 'simple_numbers', 'full', 'full_numbers', 'first_last_numbers')] $PagingStyle = 'full_numbers',
        [int[]]$PagingOptions = @(15, 25, 50, 100),
        [switch]$DisablePaging,
        [switch]$DisableOrdering,
        [switch]$DisableInfo,
        [switch]$HideFooter,
        [switch]$DisableColumnReorder,
        [switch]$DisableProcessing,
        [switch]$DisableResponsiveTable,
        [switch]$DisableSelect,
        [switch]$DisableStateSave,
        [switch]$DisableSearch,
        [switch]$ScrollCollapse,
        [switch]$OrderMulti,
        [switch]$Filtering,
        [ValidateSet('Top', 'Bottom', 'Both')][string]$FilteringLocation = 'Bottom',
        [string[]][ValidateSet('display', 'cell-border', 'compact', 'hover', 'nowrap', 'order-column', 'row-border', 'stripe')] $Style = @('display', 'compact'),
        [switch]$Simplify,
        [string]$TextWhenNoData = 'No data available.',
        [int] $ScreenSizePercent = 0,
        [string[]] $DefaultSortColumn,
        [int[]] $DefaultSortIndex,
        [ValidateSet('Ascending', 'Descending')][string] $DefaultSortOrder = 'Ascending',
        [string[]]$DateTimeSortingFormat,
        [alias('Search')][string]$Find,
        [switch] $InvokeHTMLTags,
        [switch] $DisableNewLine,
        [switch] $ScrollX,
        [switch] $ScrollY,
        [int] $ScrollSizeY = 500,
        [int] $FreezeColumnsLeft,
        [int] $FreezeColumnsRight,
        [switch] $FixedHeader,
        [switch] $FixedFooter,
        [string[]] $ResponsivePriorityOrder,
        [int[]] $ResponsivePriorityOrderIndex,
        [string[]] $PriorityProperties,
        [switch] $ImmediatelyShowHiddenDetails,
        [alias('RemoveShowButton')][switch] $HideShowButton,
        [switch] $AllProperties,
        [switch] $SkipProperties,
        [switch] $Compare,
        [alias('CompareWithColors')][switch] $HighlightDifferences,
        [int] $First,
        [int] $Last,
        [alias('Replace')][Array] $CompareReplace
    )
    Begin {
        $DataTable = [System.Collections.Generic.List[Object]]::new()
        if ($FilePath -eq '') {
            $FilePath = Get-FileName -Extension 'html' -Temporary
        }
    }
    Process {
        if ($null -ne $Table) {
            foreach ($T in $Table) {
                $DataTable.Add($T)
            }
        }
    }
    End {
        if ($null -ne $Table) {
            # HTML generation part
            New-HTML -FilePath $FilePath -UseCssLinks -UseJavaScriptLinks -TitleText $Title -ShowHTML {
                New-HTMLTable -DataTable $DataTable `
                    -HideFooter:$HideFooter `
                    -Buttons $Buttons -PagingStyle $PagingStyle -PagingOptions $PagingOptions `
                    -DisablePaging:$DisablePaging -DisableOrdering:$DisableOrdering -DisableInfo:$DisableInfo -DisableColumnReorder:$DisableColumnReorder -DisableProcessing:$DisableProcessing `
                    -DisableResponsiveTable:$DisableResponsiveTable -DisableSelect:$DisableSelect -DisableStateSave:$DisableStateSave -DisableSearch:$DisableSearch -ScrollCollapse:$ScrollCollapse `
                    -Style $Style -TextWhenNoData:$TextWhenNoData -ScreenSizePercent $ScreenSizePercent `
                    -HTML $HTML -PreContent $PreContent -PostContent $PostContent `
                    -DefaultSortColumn $DefaultSortColumn -DefaultSortIndex $DefaultSortIndex -DefaultSortOrder $DefaultSortOrder `
                    -DateTimeSortingFormat $DateTimeSortingFormat -Find $Find -OrderMulti:$OrderMulti `
                    -Filtering:$Filtering -FilteringLocation $FilteringLocation `
                    -InvokeHTMLTags:$InvokeHTMLTags -DisableNewLine:$DisableNewLine -ScrollX:$ScrollX -ScrollY:$ScrollY -ScrollSizeY $ScrollSizeY `
                    -FreezeColumnsLeft $FreezeColumnsLeft -FreezeColumnsRight $FreezeColumnsRight `
                    -FixedHeader:$FixedHeader -FixedFooter:$FixedFooter -ResponsivePriorityOrder $ResponsivePriorityOrder `
                    -ResponsivePriorityOrderIndex $ResponsivePriorityOrderIndex -PriorityProperties $PriorityProperties -AllProperties:$AllProperties `
                    -SkipProperties:$SkipProperties -Compare:$Compare -HighlightDifferences:$HighlightDifferences -First $First -Last $Last `
                    -ImmediatelyShowHiddenDetails:$ImmediatelyShowHiddenDetails -Simplify:$Simplify -HideShowButton:$HideShowButton -CompareReplace $CompareReplace
            }
            if ($PassThru) {
                # This isn't really real PassThru but just passing final object further down the pipe when needed
                # real PassThru requires significant work - if you're up to it, let me know.
                $DataTable
            }
        } else {
            Write-Warning 'Out-HtmlView - No data available.'
        }
    }
}

Function Save-HTML {
    <#
    .SYNOPSIS
    #

    .DESCRIPTION
    Long description

    .PARAMETER FilePath
    Parameter description

    .PARAMETER HTML
    Parameter description

    .PARAMETER ShowHTML
    Parameter description

    .PARAMETER Encoding
    Parameter description

    .PARAMETER Supress
    Parameter description

    .EXAMPLE
    An example

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $false)][string]$FilePath,
        [Parameter(Mandatory = $true)][Array] $HTML,
        [alias('Show', 'Open')][Parameter(Mandatory = $false)][switch]$ShowHTML,
        [ValidateSet('Unknown', 'String', 'Unicode', 'Byte', 'BigEndianUnicode', 'UTF8', 'UTF7', 'UTF32', 'Ascii', 'Default', 'Oem', 'BigEndianUTF32')] $Encoding = 'UTF8',
        [bool] $Supress = $true
    )
    if ([string]::IsNullOrEmpty($FilePath)) {
        $FilePath = Get-FileName -Temporary -Extension 'html'
        Write-Verbose "Save-HTML - FilePath parameter is empty, using Temporary $FilePath"
    } else {
        if (Test-Path -LiteralPath $FilePath) {
            Write-Verbose "Save-HTML - Path $FilePath already exists. Report will be overwritten."
        }
    }
    Write-Verbose "Save-HTML - Saving HTML to file $FilePath"
    try {
        $HTML | Set-Content -LiteralPath $FilePath -Force -Encoding $Encoding -ErrorAction Stop
        if (-not $Supress) {
            $FilePath
        }
    } catch {
        $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
        $FilePath = Get-FileName -Temporary -Extension 'html'
        Write-Warning "Save-HTML - Failed with error: $ErrorMessage"
        Write-Warning "Save-HTML - Saving HTML to file $FilePath"
        try {
            $HTML | Set-Content -LiteralPath $FilePath -Force -Encoding $Encoding -ErrorAction Stop
        } catch {
            $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
            Write-Warning "Save-HTML - Failed with error: $ErrorMessage`nPlease define a different path for the `'-FilePath`' parameter."
        }
    }
    if ($ShowHTML) {
        try {
            Invoke-Item -LiteralPath $FilePath -ErrorAction Stop
        } catch {
            $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
            Write-Verbose "Save-HTML - couldn't open file $FilePath in a browser. Error: $ErrorMessage"
        }
    }
}



Export-ModuleMember -Function @('Email', 'EmailAttachment', 'EmailBCC', 'EmailBody', 'EmailCC', 'EmailFrom', 'EmailHeader', 'EmailHTML', 'EmailListItem', 'EmailOptions', 'EmailReplyTo', 'EmailServer', 'EmailSubject', 'EmailText', 'EmailTextBox', 'EmailTo', 'New-CalendarEvent', 'New-ChartAxisX', 'New-ChartAxisY', 'New-ChartBar', 'New-ChartBarOptions', 'New-ChartDonut', 'New-ChartGrid', 'New-ChartLegend', 'New-ChartLine', 'New-ChartPie', 'New-ChartRadial', 'New-ChartTheme', 'New-ChartToolbar', 'New-DiagramEvent', 'New-DiagramLink', 'New-DiagramNode', 'New-DiagramOptionsInteraction', 'New-DiagramOptionsLayout', 'New-DiagramOptionsLinks', 'New-DiagramOptionsManipulation', 'New-DiagramOptionsNodes', 'New-DiagramOptionsPhysics', 'New-GageSector', 'New-HierarchicalTreeNode', 'New-HTML', 'New-HTMLCalendar', 'New-HTMLChart', 'New-HTMLCodeBlock', 'New-HTMLContainer', 'New-HTMLDiagram', 'New-HTMLFooter', 'New-HTMLGage', 'New-HTMLHeader', 'New-HTMLHeading', 'New-HTMLHierarchicalTree', 'New-HTMLHorizontalLine', 'New-HTMLImage', 'New-HTMLList', 'New-HTMLListItem', 'New-HTMLLogo', 'New-HTMLMain', 'New-HTMLPanel', 'New-HTMLSection', 'New-HTMLSpanStyle', 'New-HTMLStatus', 'New-HTMLStatusItem', 'New-HTMLTab', 'New-HTMLTable', 'New-HTMLTabOptions', 'New-HTMLTag', 'New-HTMLText', 'New-HTMLTimeline', 'New-HTMLTimelineItem', 'New-HTMLToast', 'New-HTMLTree', 'New-TableButtonCopy', 'New-TableButtonCSV', 'New-TableButtonExcel', 'New-TableButtonPageLength', 'New-TableButtonPDF', 'New-TableButtonPrint', 'New-TableCondition', 'New-TableContent', 'New-TableHeader', 'New-TableReplace', 'New-TableRowGrouping', 'New-TreeNode', 'Out-HtmlView', 'Save-HTML') -Alias @('Calendar', 'CalendarEvent', 'Chart', 'ChartAxisX', 'ChartAxisY', 'ChartBar', 'ChartBarOptions', 'ChartCategory', 'ChartDonut', 'ChartGrid', 'ChartLegend', 'ChartLine', 'ChartPie', 'ChartRadial', 'ChartTheme', 'ChartToolbar', 'Container', 'Dashboard', 'Diagram', 'DiagramEdge', 'DiagramEdges', 'DiagramLink', 'DiagramNode', 'DiagramOptionsEdges', 'DiagramOptionsInteraction', 'DiagramOptionsLayout', 'DiagramOptionsLinks', 'DiagramOptionsManipulation', 'DiagramOptionsNodes', 'DiagramOptionsPhysics', 'EmailList', 'EmailTable', 'EmailTableButtonCopy', 'EmailTableButtonCSV', 'EmailTableButtonExcel', 'EmailTableButtonPageLength', 'EmailTableButtonPDF', 'EmailTableButtonPrint', 'EmailTableCondition', 'EmailTableContent', 'EmailTableHeader', 'EmailTableReplace', 'EmailTableRowGrouping', 'Footer', 'Header', 'HierarchicalTreeNode', 'HTMLText', 'Image', 'Main', 'New-ChartCategory', 'New-DiagramEdge', 'New-DiagramOptionsEdges', 'New-HierarchicalTreeNode', 'New-HTMLColumn', 'New-HTMLContent', 'New-HTMLTableButtonCopy', 'New-HTMLTableButtonCSV', 'New-HTMLTableButtonExcel', 'New-HTMLTableButtonPageLength', 'New-HTMLTableButtonPDF', 'New-HTMLTableButtonPrint', 'New-HTMLTableCondition', 'New-HTMLTableContent', 'New-HTMLTableHeader', 'New-HTMLTableReplace', 'New-HTMLTableRowGrouping', 'ohv', 'Out-GridHtml', 'Panel', 'Section', 'Tab', 'Table', 'TableButtonCopy', 'TableButtonCSV', 'TableButtonExcel', 'TableButtonPageLength', 'TableButtonPDF', 'TableButtonPrint', 'TableCondition', 'TableConditionalFormatting', 'TableContent', 'TableHeader', 'TableReplace', 'TableRowGrouping', 'TabOptions', 'Text')