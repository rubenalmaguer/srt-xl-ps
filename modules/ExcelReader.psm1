# ExcelReader.psm1

# Big assumption! Don't mix these two up: 29.976 vs. 23.96
# Used when reading frame-based timestamps (00;00;00;00 or 00:00:00;00)
# IRRELEVANT for milliseconds-based timestamps (00:00:00,000)
$FRAME_RATE = 29.976

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

    # Handle various Excel time formats (e.g. 00:00:00,00 <- Seen this as a form of cropped milliseconds)
    if ($TimeValue -match '^(\d{1,2}):(\d{2}):(\d{2})[\.,](\d{1,3})$') {
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

        # Excel serial dates are typically between 1 and 50000+ (year 2136)
        # For time-only values, should be between 0 and 1
        if ($serialTime -gt 1000) {
            Write-Warning "Value $TimeValue seems too large for Excel serial time format, skipping this format"
            return [int]0  # TODO: Handle appropriately
        }

        $totalMilliseconds = $serialTime * 24 * 60 * 60 * 1000

        # Convert to long first to avoid overflow, then check if it fits in int
        $totalMs = [long]$totalMilliseconds
    
        if ($totalMs -gt [int]::MaxValue -or $totalMs -lt [int]::MinValue) {
            Write-Host "Result too large for Int32, returning as Long: $totalMs"
            return $totalMs  # Return as long
        }
        return [int]$totalMs
    }
    elseif ($TimeValue -match '^(\d{1,2})[:;](\d{2})[:;](\d{2})[:;](\d{2})$') {
        
        # Frames-based Format

        $hours = $matches[1]
        $minutes = $matches[2]
        $seconds = $matches[3]
        $frames = $matches[4] #.PadLeft(2, '0')

        <#         Write-Host "-------------------------------------------"
        Write-Host "Value: $TimeValue"
        Write-Host "Frames: $frames" #>

        $milliseconds = [int]([int]$frames * (1000 / $FRAME_RATE))

        <# Write-Host "Milliseconds: $milliseconds" #>

        # If you need it in reverse: $frames = [int]$milliseconds * ($FRAME_RATE / 1000)

        return ([int]$hours * 3600000) + ([int]$minutes * 60000) + ([int]$seconds * 1000) + [int]$milliseconds

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

        [Parameter(Mandatory = $false)]
        [string]$SheetName = "Subtitles"  # Optional: default sheet name
    )
    try {
        # Create Excel COM object
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

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
            $startTime = Convert-MillisecondsToTimestamp -Milliseconds $cue.startMS -FrameRate $FRAME_RATE

            $endTime = Convert-MillisecondsToTimestamp -Milliseconds $cue.endMS -FrameRate $FRAME_RATE

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
        [int]$Milliseconds,

        [Parameter(Mandatory = $false)]
        [int]$FrameRate = 0 # If provided, use 00;00;00;00 format (with frames) instead of 00:00:00,000 (with milliseconds)
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

    $formattedTimestamp = "$formattedHours`:$formattedMinutes`:$formattedSeconds,$formattedMS"

    if ($FrameRate -gt 0) {
        $frames = [math]::Round($ms * ($FrameRate / 1000))
        $formattedFRAMES = $frames.ToString("00")
        $formattedTimestamp = "$formattedHours`;$formattedMinutes`;$formattedSeconds`;$formattedFRAMES"
    }

    return $formattedTimestamp
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
Export-ModuleMember -Function Convert-CuesToExcel, Convert-ExcelToCues