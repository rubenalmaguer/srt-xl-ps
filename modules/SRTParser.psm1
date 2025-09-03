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
    $Content | Out-File -LiteralPath $FilePath -Encoding $Encoding
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
Export-ModuleMember -Function Convert-SrtToCues, Convert-CuesToSrt, ConvertTo-Milliseconds, ConvertFrom-Milliseconds, Out-FileSafe