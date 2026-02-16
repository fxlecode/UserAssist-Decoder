function Get-UserAssist {
    [CmdletBinding()]
    param(
        [switch]$ExportCSV,
        [string]$OutputPath = "UserAssist_Report.csv",
        [switch]$IncludeRawData
    )

    $KnownFolders = @{
        "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}" = "%SystemRoot%\System32"
        "{6D809371-2109-4E85-AE99-FD3C73560747}" = "%ProgramFiles%"
        "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}" = "%ProgramFiles(x86)%"
        "{F38BF404-1D43-42F2-9305-67DE0B28FC23}" = "%UserProfile%"
        "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}" = "%SystemRoot%\SysWOW64"
        "{905E63B6-C1BF-494E-B29C-65B732D3D21A}" = "%ProgramData%"
        "{5E6C858F-0E22-4760-9AFE-EA3317B67173}" = "%UserProfile%\Downloads"
        "{F42EE2D3-909F-4907-8871-4C22FC0BF756}" = "%UserProfile%\Documents"
        "{0DDD015D-B06C-45D5-8C4C-F59713854639}" = "%UserProfile%\Pictures"
        "{35286A68-3C57-41A1-BBB1-0EAE73D76C95}" = "%UserProfile%\Videos"
        "{A0C69A99-21C8-4671-8703-7934162FCF1D}" = "%UserProfile%\Music"
        "{4BFEFB45-347D-4006-A5BE-AC0CB0567192}" = "%UserProfile%\Desktop"
        "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}" = "%UserProfile%\Desktop"
        "{374DE290-123F-4565-9164-39C4925E467B}" = "%UserProfile%\Downloads"
        "{56784854-C6CB-462B-8169-88E350ACB882}" = "%UserProfile%\Contacts"
        "{1777F761-68AD-4D8A-87BD-30B759FA33DD}" = "%UserProfile%\Favorites"
        "{3D644C9B-1FB8-4F30-9B45-F670235F79C0}" = "%UserProfile%\Saved Games"
        "{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}" = "%UserProfile%\Searches"
        "{9E3995AB-1F9C-4F13-B827-48B24B6C7174}" = "%UserProfile%\OneDrive"
        "{5CD7AEE2-2219-4A67-B85D-6C9CE15660CB}" = "%UserProfile%\OneDrive\Pictures"
        "{31C0DD25-9439-4F12-BF41-7FF4EDA38722}" = "%UserProfile%\3D Objects"
    }

    $UserAssistGUIDs = @(
        "{CEBFF5CD-ACE2-4F4F-9178-9926F41749EA}",
        "{F4E57140-A759-4561-892D-1991F0467B03}"
    )

    $Results = New-Object System.Collections.ArrayList

    foreach ($GUID in $UserAssistGUIDs) {
        $RegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist\$GUID\Count"
        
        if (-not (Test-Path $RegistryPath)) {
            Write-Warning "Registry path not found: $GUID"
            continue
        }

        $RegistryItem = Get-Item -Path $RegistryPath -ErrorAction SilentlyContinue
        $Properties = Get-ItemProperty -Path $RegistryPath -ErrorAction SilentlyContinue

        foreach ($ValueName in $RegistryItem.Property) {
            $DecodedPath = Invoke-ROT13Decode -InputString $ValueName
            $ResolvedPath = Resolve-KnownFolderGUIDs -InputPath $DecodedPath -LookupTable $KnownFolders
            $BinaryData = $Properties.$ValueName

            if ($BinaryData.Length -lt 72) { continue }

            $RawCount = [BitConverter]::ToInt32($BinaryData, 4)
            $AdjustedCount = if ($RawCount -ge 5) { $RawCount - 5 } else { $RawCount }
            
            $FileTimeStamp = [BitConverter]::ToInt64($BinaryData, 60)
            $LastExecution = if ($FileTimeStamp -gt 0) { 
                [DateTime]::FromFileTime($FileTimeStamp) 
            } else { 
                [DateTime]::MinValue 
            }

            $FirstExecution = if ($BinaryData.Length -ge 80 -and [BitConverter]::ToInt64($BinaryData, 68) -gt 0) { 
                [DateTime]::FromFileTime([BitConverter]::ToInt64($BinaryData, 68))
            } else { 
                $null
            }

            $FocusTimeMs = if ($BinaryData.Length -ge 16) {
                [BitConverter]::ToInt32($BinaryData, 12)
            } else {
                0
            }
            
            $FocusTimeFormatted = Format-FocusTime -Milliseconds $FocusTimeMs

            $EvidenceEntry = [PSCustomObject]@{
                ProgramPath       = $ResolvedPath
                RawPath           = if ($IncludeRawData) { $DecodedPath } else { $null }
                ExecutionCount    = $AdjustedCount
                RawCount          = if ($IncludeRawData) { $RawCount } else { $null }
                CountBiasApplied  = ($RawCount -ne $AdjustedCount)
                LastExecuted      = $LastExecution
                FirstExecuted     = $FirstExecution
                FocusTimeMs       = $FocusTimeMs
                FocusTimeReadable = $FocusTimeFormatted
                RegistrySource    = $GUID
                LastWriteTime     = $RegistryItem.LastWriteTime
            }

            if (-not $IncludeRawData) {
                $EvidenceEntry.PSObject.Properties.Remove('RawPath')
                $EvidenceEntry.PSObject.Properties.Remove('RawCount')
            }

            [void]$Results.Add($EvidenceEntry)
        }
    }

    $SortedResults = $Results | Sort-Object LastExecuted -Descending

    if ($ExportCSV) {
        $SortedResults | Select-Object * | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "[+] Report exported to: $OutputPath" -ForegroundColor Green
    }

    return $SortedResults
}

function Invoke-ROT13Decode {
    param([string]$InputString)
    
    $DecodedChars = $InputString.ToCharArray() | ForEach-Object {
        $CharCode = [int]$_
        
        if ($CharCode -ge 65 -and $CharCode -le 90) {
            [char]((($CharCode - 65 + 13) % 26) + 65)
        } elseif ($CharCode -ge 97 -and $CharCode -le 122) {
            [char]((($CharCode - 97 + 13) % 26) + 97)
        } else {
            $_
        }
    }
    
    return -join $DecodedChars
}

function Resolve-KnownFolderGUIDs {
    param(
        [string]$InputPath,
        [hashtable]$LookupTable
    )
    
    $ResolvedPath = $InputPath
    
    foreach ($GUID in $LookupTable.Keys) {
        if ($ResolvedPath -like "*$GUID*") {
            $ResolvedPath = $ResolvedPath -replace [regex]::Escape($GUID), $LookupTable[$GUID]
        }
    }
    
    return $ResolvedPath
}

function Format-FocusTime {
    param([int]$Milliseconds)
    
    if ($Milliseconds -eq 0) { return "N/A" }
    if ($Milliseconds -lt 1000) { return "$Milliseconds ms" }
    
    $TimeSpan = [TimeSpan]::FromMilliseconds($Milliseconds)
    
    if ($TimeSpan.TotalHours -ge 1) {
        return "{0:D2}h {1:D2}m {2:D2}s" -f $TimeSpan.Hours, $TimeSpan.Minutes, $TimeSpan.Seconds
    } elseif ($TimeSpan.TotalMinutes -ge 1) {
        return "{0:D2}m {1:D2}s" -f $TimeSpan.Minutes, $TimeSpan.Seconds
    } else {
        return "{0:D2}s" -f $TimeSpan.Seconds
    }
}

function Show-UserAssistReport {
    param([switch]$ShowAll)

    $Data = Get-UserAssist -IncludeRawData:$ShowAll
    
    if (-not $Data) {
        Write-Host "[!] No UserAssist data found." -ForegroundColor Red
        return
    }

    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "              USERASSIST FORENSIC DECODER - EXECUTION EVIDENCE" -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host ""

    $Data | Format-Table -Property @{
        Name = "Program"; Expression = { 
            $MaxLen = 45
            if ($_.ProgramPath.Length -gt $MaxLen) {
                $_.ProgramPath.Substring(0, $MaxLen) + "..."
            } else {
                $_.ProgramPath
            }
        }; Width = 48
    }, @{
        Name = "Runs"; Expression = { $_.ExecutionCount }; Width = 6; Align = "Right"
    }, @{
        Name = "Focus Time"; Expression = { $_.FocusTimeReadable }; Width = 14
    }, @{
        Name = "Last Execution"; Expression = { 
            if ($_.LastExecuted -eq [DateTime]::MinValue) { "Never" } 
            else { $_.LastExecuted.ToString("yyyy-MM-dd HH:mm") }
        }; Width = 18
    }, @{
        Name = "Bias"; Expression = { if ($_.CountBiasApplied) { "*" } else { "" } }; Width = 6; Align = "Center"
    } -AutoSize

    $TotalFocusTime = ($Data | Measure-Object -Property FocusTimeMs -Sum).Sum
    $ActivePrograms = ($Data | Where-Object { $_.FocusTimeMs -gt 0 }).Count

    Write-Host "================================================================================" -ForegroundColor DarkGray
    Write-Host "  STATISTICS:" -ForegroundColor Yellow
    Write-Host "  * Total entries decoded: $($Data.Count)" -ForegroundColor White
    Write-Host "  * Programs with active usage: $ActivePrograms" -ForegroundColor White
    Write-Host "  * Total system focus time: $(Format-FocusTime -Milliseconds $TotalFocusTime)" -ForegroundColor White
    Write-Host "  * Most recent execution: $($Data[0].LastExecuted)" -ForegroundColor White
    if (($Data | Where-Object { $_.CountBiasApplied }).Count -gt 0) {
        Write-Host "  * (*) Indicates Windows bias correction applied (Count - 5)" -ForegroundColor DarkGray
    }
    Write-Host "================================================================================" -ForegroundColor DarkGray
    Write-Host ""
}

Show-UserAssistReport