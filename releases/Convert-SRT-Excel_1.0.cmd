@@ echo off
@@ title SRT-EXCEL CONVERTER
@@ chcp 65001 >NUL
@@ echo Starting...
@@ REM echo.
@@ setlocal
@@ set PS_WRAPPER_ARGS=%*
@@ set PS_WRAPPER_PATH=%~f0
@@ if defined PS_WRAPPER_ARGS set PS_WRAPPER_ARGS=%PS_WRAPPER_ARGS:"=\"%
@@ PowerShell -NoExit -sta -Command Invoke-Expression $('$args=@(^&{$args} %PS_WRAPPER_ARGS%);'+[String]::Join([Environment]::NewLine,$((Get-Content '%PS_WRAPPER_PATH%') -notmatch '^^@@^|^^:^|^^cls'))) & endlocal & exit
{
####################### DRAG-N-DROP WRAPPER #######################
#Requires -version 2.0
Set-StrictMode -version 2.0
$ErrorActionPreference = "stop"
$script:selfPath = [environment]::GetEnvironmentVariable("PS_WRAPPER_PATH")
$script:cmdPid = (gwmi win32_process -Filter "processid='$pid'").parentprocessid
# To access args/dragged files = $script:args (e.g. $script:args[1])

######################### EMBEDDED SCRIPT START ###########################

#Region ExcelReader Functions
# ExcelReader Functions
# ExcelReader.psm1
function Convert-ExcelToCues {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExcelPath,
        [string]$SheetName = "", # Optional, defaults to first sheet if empty
        [int]$DefaultDurationSeconds = 2  # Default duration when end time is missing
    )
    if (-not $(Test-Path -LiteralPath $ExcelPath)) {
        return @("Path not found.", $null)
    }
    try {
        # Create Excel COM object
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        # Open the workbook
        $workbook = $excel.Workbooks.Open($ExcelPath)
        # Get the first worksheet if no sheet name provided
        if ($SheetName -eq "") {
            $worksheet = $workbook.Worksheets.Item(1)
        }
        else {
            $worksheet = $workbook.Worksheets.Item($SheetName)
        }
        # Find the last used row
        $lastRow = $worksheet.UsedRange.Rows.Count
        # Initialize cues array
        $cues = @()
        # Start from row 2 to skip header
        for ($row = 2; $row -le $lastRow; $row++) {
            # Read cell values
            $startTime = $worksheet.Cells($row, 1).Text
            $endTime = $worksheet.Cells($row, 2).Text
            $text = $worksheet.Cells($row, 3).Text
            # Skip if start time or text is empty
            if ([string]::IsNullOrWhiteSpace($startTime) -or 
                [string]::IsNullOrWhiteSpace($text)) {
                continue
            }
            # Convert start time
            $startMS = Convert-ExcelTimeToMilliseconds -TimeValue $startTime
            # If end time is missing, use default duration
            if ([string]::IsNullOrWhiteSpace($endTime)) {
                $endMS = $startMS + ($DefaultDurationSeconds * 1000)
            }
            else {
                $endMS = Convert-ExcelTimeToMilliseconds -TimeValue $endTime
            }
            # Create cue object
            $cue = @{
                id      = $cues.Count + 1  # 1-based index
                startMS = $startMS
                endMS   = $endMS
                text    = $text.Trim()
            }
            $cues += $cue
        }
        if ($cues.length -lt 1) {
            return @("No subtitle cues found in `"$ExcelPath`"", $null)
        }
        else {
            return @($null, $cues)
        }
    }
    catch {
        Write-Error "Error processing Excel file: $_"
        throw
    }
    finally {
        # Clean up Excel objects
        if ($workbook) {
            $workbook.Close($false)
        }
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}
function Convert-ExcelTimeToMilliseconds {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TimeValue
    )
    # Check if the time is already in SRT format (00:00:00,000)
    if ($TimeValue -match '^\d{2}:\d{2}:\d{2},\d{3}$') {
        return ConvertTo-Milliseconds -TimeStamp $TimeValue
    }
    # Handle various Excel time formats
    if ($TimeValue -match '^(\d{1,2}):(\d{2}):(\d{2})\.(\d{1,3})$') {
        # Format: HH:MM:SS.mmm
        $hours = $matches[1]
        $minutes = $matches[2]
        $seconds = $matches[3]
        $milliseconds = $matches[4].PadRight(3, '0')
        return ([int]$hours * 3600000) + ([int]$minutes * 60000) + ([int]$seconds * 1000) + [int]$milliseconds
    }
    elseif ($TimeValue -match '^(\d{1,2}):(\d{2})$') {
        # Format: MM:SS
        $minutes = $matches[1]
        $seconds = $matches[2]
        return ([int]$minutes * 60000) + ([int]$seconds * 1000)
    }
    elseif ($TimeValue -match '^\d+\.?\d*$') {
        # Excel serial time format (fraction of 24 hours)
        $serialTime = [double]$TimeValue
        $totalMilliseconds = $serialTime * 24 * 60 * 60 * 1000
        return [int]$totalMilliseconds
    }
    else {
        throw "Unsupported time format: $TimeValue"
    }
}
function Convert-CuesToExcel {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Cues, # Array of cues to write to Excel
        [Parameter(Mandatory = $true)]
        [string]$ExcelPath, # Desired path of the output Excel file
        [string]$SheetName = "Subtitles"  # Optional: default sheet name
    )
    try {
        # Create Excel COM object
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $true
        # Create a new workbook
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = $SheetName
        # Set headers
        $worksheet.Cells(1, 1).Value = "Start Time"
        $worksheet.Cells(1, 2).Value = "End Time"
        $worksheet.Cells(1, 3).Value = "Text"
        # Start writing from row 2
        $row = 2
        foreach ($cue in $Cues) {
            # Convert start and end times from milliseconds to timestamp string
            $startTime = Convert-MillisecondsToTimestamp -Milliseconds $cue.startMS
            $endTime = Convert-MillisecondsToTimestamp -Milliseconds $cue.endMS
            # Write data to Excel
            $worksheet.Cells($row, 1).Value = $startTime
            $worksheet.Cells($row, 2).Value = $endTime
            $worksheet.Cells($row, 3).Value = $cue.text
            $row++
        }
        # Auto-fit columns
        $worksheet.Columns.AutoFit()
        # Wider Text column (C [3])
        $column3 = $worksheet.Columns.Item(3)
        $column3.ColumnWidth = $column3.ColumnWidth * 5
        # Save
        $savedPath = Save-ExcelSafe -Workbook $workbook -Path $ExcelPath
        if ($savedPath) {
            Write-Host "Saved file: $savedPath" -BackgroundColor Green -ForegroundColor Black
        }
        else {
            Write-Host "Failed to save file: $ExcelPath" -ForegroundColor Red
        }
    }
    catch {
        Write-Error "Error processing Excel file: $_"
        throw
    }
    finally {
        # Clean up Excel objects
        <#         if (-not $null -eq $workbook) {
            $workbook.Close($false)
        } #>
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
            if ($workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null }
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}
function Convert-MillisecondsToTimestamp {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Milliseconds
    )
    $MS_PER_HR = 3600000
    $MS_PER_MIN = 60000
    $MS_PER_SECOND = 1000
    $hours = [Math]::Floor($Milliseconds / $MS_PER_HR)
    $remainder = $Milliseconds % $MS_PER_HR
    $minutes = [Math]::Floor($remainder / $MS_PER_MIN)
    $remainder = $remainder % $MS_PER_MIN
    $seconds = [Math]::Floor($remainder / $MS_PER_SECOND)
    $ms = $remainder % $MS_PER_SECOND
    $formattedHours = $hours.ToString("00")
    $formattedMinutes = $minutes.ToString("00")
    $formattedSeconds = $seconds.ToString("00")
    $formattedMS = $ms.ToString("000")
    return "$formattedHours`:$formattedMinutes`:$formattedSeconds,$formattedMS"
}
function Save-ExcelSafe {
    <#
    .SYNOPSIS
    Excel file name/paths CANNOT contain brackets, even though Windows allows it!
    This can be circumvented via the file system,
    which happens inadvertently when converting from a different format
    and is usually not a problem, except for in the edge case
    where the brackets are the only difference, e.g. Example.xlsx and [Example].xlsx
    In that case, double clicking [Example].xlsx opens Example.xlsx.
    As far as i can tell, the edge case only affects file names, not file paths.
    However, Excel will still refuse to save to a path containing brackets.
    SOLUTION:
    Save in temp path, then move to originally requested path, but with safe file name
    WARNING:
    Don't use Workbook after calling this function
    #>
    param (
        [Parameter(Mandatory = $true)]
        [Object]$Workbook,
        [string]$Path, # Desired output path
        [string]$BracketReplacement = ""
    )
    $Directory = Split-Path -Path $Path -Parent
    $FileName = Split-Path -Path $Path -Leaf
    # Remove brackets
    $SafeFileName = $FileName -replace '[\[\]<>\?|]', $BracketReplacement
    # Save in safe, temp folder (Assuming TEMP path (User name) contains no brackets)
    $TempPath = Join-Path -Path $env:TEMP -ChildPath $SafeFileName
    for ($i = 2; Test-Path -LiteralPath $TempPath; $i++) { $TempPath = $TempPath -replace '(?:_\d+)?\.(\w+)$', "_$i.`$1" } # Ensure uniqueness
    try {
        $Workbook.SaveAs($TempPath)
        $Workbook.Close($false)
    }
    catch {
        write-host Error saving temp file: $_
        Write-host $TempPath
        return $null
    }
    # Create output directory if it doesn't exist
    if ($Directory -and -not (Test-Path -LiteralPath $Directory)) {
        New-Item -ItemType Directory -Path $Directory -Force | Out-Null
    }
    # Move to requested location
    $SafePath = Join-Path -Path $Directory -ChildPath $SafeFileName
    for ($i = 2; Test-Path -LiteralPath $SafePath; $i++) { $SafePath = $SafePath -replace '(?:_\d+)?\.(\w+)$', "_$i.`$1" } # Ensure uniqueness
    try {
        Move-Item -Path $TempPath -Destination $SafePath
    }
    catch {
        write-host Error moving file: $_
        Write-host $TempPath
        Write-host $SafePath
        return $null
    }
    return $SafePath
}
# Export functions


#EndRegion ExcelReader Functions
#Region SRTParser Functions
# SRTParser Functions
# SRTParser.psm1
# Define regex patterns
$script:Patterns = @{
    LineBreak              = "(?:\s*\r?\n\s*)+"
    MaybeLineBreak         = "(?:\s*\r?\n\s*)*"
    Id                     = "(\d+)"
    MaybeHoursAndMinutes   = "(?:\d{1,3}:){0,2}"
    SecondsAndMilliseconds = "\d{1,3}[,\.]\d{1,3}"
    Arrow                  = " *--> *"
}
# Build the complete regex pattern
$script:TimecodeRegex = $null
function Initialize-TimecodeRegex {
    $singleTimestamp = "($($Patterns.MaybeHoursAndMinutes)$($Patterns.SecondsAndMilliseconds))"
    $fullPattern = "$($Patterns.MaybeLineBreak)" +
    "$($Patterns.Id)" +
    "$($Patterns.LineBreak)" +
    "$singleTimestamp" +
    "$($Patterns.Arrow)" +
    "$singleTimestamp" +
    "$($Patterns.LineBreak)"
    $script:TimecodeRegex = $fullPattern
}
function ConvertTo-Milliseconds {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TimeStamp
    )
    $result = 0
    $parts = $TimeStamp -split '[,\.]'
    $ms = $parts[1].PadRight(3, '0')
    $result += [int]$ms
    $hms = $parts[0] -split ':'
    $seconds = if ($hms.Length -gt 0) { $hms[-1] } else { "0" }
    $result += [int]$seconds * 1000
    $minutes = if ($hms.Length -gt 1) { $hms[-2] } else { "0" }
    $result += [int]$minutes * 60 * 1000
    $hours = if ($hms.Length -gt 2) { $hms[-3] } else { "0" }
    $result += [int]$hours * 60 * 60 * 1000
    return $result
}
function ConvertFrom-Milliseconds {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Milliseconds
    )
    $MS_PER_HR = 3600000
    $MS_PER_MIN = 60000
    $MS_PER_SECOND = 1000
    $hours = [Math]::Floor($Milliseconds / $MS_PER_HR)
    $remainder = $Milliseconds % $MS_PER_HR
    $minutes = [Math]::Floor($remainder / $MS_PER_MIN)
    $remainder = $remainder % $MS_PER_MIN
    $seconds = [Math]::Floor($remainder / $MS_PER_SECOND)
    $ms = $remainder % $MS_PER_SECOND
    $formattedHours = $hours.ToString("00")
    $formattedMinutes = $minutes.ToString("00")
    $formattedSeconds = $seconds.ToString("00")
    $formattedMS = $ms.ToString("000")
    return "$formattedHours`:$formattedMinutes`:$formattedSeconds,$formattedMS"
}
function Out-FileSafe {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$Content,
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [Parameter(Mandatory = $false)]
        [ValidateSet('utf8', 'unicode', 'bigendianunicode', 'ascii', 'default', 'oem')]
        [string]$Encoding = 'utf8'
    )
    # Get the directory path
    $directory = Split-Path -Parent $FilePath
    # Create directory if it doesn't exist
    if ($directory -and -not (Test-Path -Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    # Write the file
    $Content | Out-File -FilePath $FilePath -Encoding $Encoding
}
function Convert-SrtToCues {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputString
    )
    # Initialize regex if not already done
    if (-not $script:TimecodeRegex) {
        Initialize-TimecodeRegex
    }
    # Remove BOM and trim whitespace
    $InputString = $InputString -replace '^\uFEFF|\uFFFE', ''
    $InputString = $InputString.Trim()
    # Split the input using regex
    $rxMatches = [regex]::Matches($InputString, $script:TimecodeRegex, [System.Text.RegularExpressions.RegexOptions]::Multiline)
    $cues = @()
    $currentIndex = 0
    foreach ($match in $rxMatches) {
        $id = $match.Groups[1].Value
        $startTime = $match.Groups[2].Value
        $endTime = $match.Groups[3].Value
        # Get text by finding the position after this match and before the next one
        $textStart = $match.Index + $match.Length
        $textEnd = if ($currentIndex -lt $rxMatches.Count - 1) {
            $rxMatches[$currentIndex + 1].Index
        }
        else {
            $InputString.Length
        }
        $text = $InputString.Substring($textStart, $textEnd - $textStart).Trim()
        $cue = @{
            id      = $id
            startMS = ConvertTo-Milliseconds -TimeStamp $startTime
            endMS   = ConvertTo-Milliseconds -TimeStamp $endTime
            text    = $text
        }
        $cues += $cue
        $currentIndex++
    }
    return $cues
}
function Convert-CuesToSrt {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Cues,
        [Parameter(Mandatory = $false)]
        [string]$OutputPath,
        [Parameter(Mandatory = $false)]
        [ValidateSet('utf8', 'unicode', 'bigendianunicode', 'ascii', 'default', 'oem')]
        [string]$Encoding = 'utf8'
    )
    $output = ""
    foreach ($cue in $Cues) {
        $startTime = ConvertFrom-Milliseconds -Milliseconds $cue.startMS
        $endTime = ConvertFrom-Milliseconds -Milliseconds $cue.endMS
        $output += "$($cue.id)`n$startTime --> $endTime`n$($cue.text)`n`n"
    }
    if ($OutputPath) {
        for ($i = 2; Test-Path -LiteralPath $OutputPath; $i++) { $OutputPath = $OutputPath -replace '(?:_\d+)?\.(\w+)$', "_$i.`$1" } # Ensure uniqueness
        $output | Out-FileSafe -FilePath $OutputPath -Encoding $Encoding
        Write-Host "Saved file: $OutputPath" -BackgroundColor Green -ForegroundColor Black
    }
    return $output
}
# Export all necessary functions


#EndRegion SRTParser Functions

#Region Main Script
# Main Script Content
$ErrorActionPreference = "continue" # "stop"

$modules = @("SRTParser", "ExcelReader")

foreach ($module in $modules) {
    # Ensure modules are not cached during development
    if (Get-Module -Name $module) {
        Remove-Module -Name $module -Force
    }
    }

$MOCK_ARGS = @("W:\F\srt-xl-ps\samples\sample-from-srt.srt", "W:\F\srt-xl-ps\samples\sample-from-xl.xlsx")

try {
    # Will fail when invoked from cmd; succeed when using .ps1
    Write-Host "$($MyInvocation.MyCommand.Name)"
    $IS_DEV = $true
}
catch {
    $IS_DEV = $false
}


function main() {
    if ($script:args.Count -eq 0 -and $IS_DEV -eq $true) {
        Write-Host "Dev mode, using mock args: $MOCK_ARGS" -ForegroundColor Red
        $srtPaths, $xlPaths, $etcArgs = ParseArgs $MOCK_ARGS
    }
    else {
        $srtPaths, $xlPaths, $etcArgs = ParseArgs $script:args
    }
    
    # Guards
    if ($srtPaths.Count -lt 1 -and $xlPaths.Count -lt 1) {
        Write-Host "You must drag at least one srt or Excel file to the Convert-SRT-Excel.cmd icon (not this window)." -BackgroundColor Yellow -ForegroundColor Black
        # ConfirmExit # CMD stays open
        Exit
    }

    $i = 0
    $srtPaths | ForEach-Object {
        $i++;
        $newPath = "{0}{1}" -f $_, ".xlsx"
        Write-Host "Converting SRT $i / $($srtPaths.Count): $_";

        try {
            if (-not $(Test-Path -LiteralPath $_)) {
                Write-Host "File not found: $_" -ForegroundColor Red;
                continue
            }
            else {
                $srtContent = Get-Content -Path $_ -Raw
            }
        }
        catch {
            Write-Host "Error reading file: $srtPath"
            continue
        }

        $cues = Convert-SrtToCues -InputString $srtContent

        try {
            Convert-CuesToExcel -Cues $cues -ExcelPath $newPath | Out-Null
        }
        catch {
            Write-Error $_
        }
    }

    $i = 0
    $xlPaths | ForEach-Object {
        $i++;
        $newPath = "{0}{1}" -f $_, ".srt"
        Write-Host "Converting Excel $i / $($xlPaths.Count): $_";

        $error, $cues = Convert-ExcelToCues -ExcelPath $_

        if ($error) {
            Write-Host $error -ForegroundColor Red
        }

        if ($cues) {
            Convert-CuesToSrt -Cues $cues -OutputPath $newPath | Out-Null
        }
    }

    # ConfirmExit # CMS stays open
    Write-Host "`nDone".
    Exit
}

function ParseArgs($arguments) {
    $srtPaths = @()
    $xlPaths = @()
    $etcArgs = @()
  
    foreach ($arg in $arguments) {
        if ($arg -match "\.(xlsx|xlsm)$") {
            $xlPaths += $arg
        }
        elseif ($arg -match "\.srt$") {
            $srtPaths += $arg
        }
        else {
            $etcArgs += $arg
        }
    }
  
    return @($srtPaths, $xlPaths, $etcArgs)
}

function ConfirmExit() {
    Write-Host "`r`nPress enter to exit." -ForegroundColor Cyan
    [Console]::ReadLine() # Read-Host adds colon.
    Exit
}

main
#EndRegion Main Script

Write-Host "
Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

############################ EMBEDDED SCRIPT END ##########################
}.Invoke($args)
