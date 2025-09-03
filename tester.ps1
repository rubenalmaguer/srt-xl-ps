$ErrorActionPreference = 'Stop'

$modules = ("SRTParser", "ExcelReader")

foreach ($module in $modules) {
    if (Get-Module -Name $module) {
        Remove-Module -Name $module -Force
    }
    Import-Module "./modules/$($module).psm1"
}

# ==============================
## FRAMES XL to SRT
Write-Host
Write-Host "FRAMES to SRT" -BackgroundColor DarkCyan -ForegroundColor White

$error, $cues = Convert-ExcelToCues -ExcelPath $([System.IO.Path]::GetFullPath("./samples/sample-with-frames.xlsx"))

if ($null -ne $error) {
    write-host $error
}
else {
    Convert-CuesToSrt -Cues $cues -OutputPath $([System.IO.Path]::GetFullPath("./samples/[out]/sample-with-frames.xlsx.srt")) | Out-Null
}

Exit

# ==============================
## MODIFY SRT
Write-Host
Write-Host "MODIFY SRT" -BackgroundColor DarkCyan -ForegroundColor White

$srtPath = [System.IO.Path]::GetFullPath("./samples/sample-from-srt.srt");

try {
    if (-not $(Test-Path -LiteralPath $srtPath)) {
        Write-Host "File not found: $srtPath"
    }
    else {
        $srtContent = Get-Content -Path $srtPath -Raw
    }
}
catch {
    Write-Host "Error reading file: $srtPath"
}

if ($srtContent) {
    $cues = Convert-SrtToCues -InputString $srtContent

    $cues[1].text = "Modified subtitle"
    
    $cues += @{
        id      = "999"
        startMS = ConvertTo-Milliseconds -TimeStamp "00:00:00,001"
        endMS   = ConvertTo-Milliseconds -TimeStamp "00:00:00,002"
        text    = "THIS IS EXTRA!!!!"
    }
    
    Convert-CuesToSrt -Cues $cues -OutputPath $([System.IO.Path]::GetFullPath("./samples/[out]/modified-subtitles.srt")) | Out-Null
}


# ==============================
## EXCEL TO SRT
Write-Host
Write-Host "EXCEL TO SRT" -BackgroundColor DarkCyan -ForegroundColor White

$error, $cues = Convert-ExcelToCues -ExcelPath $([System.IO.Path]::GetFullPath("./samples/sample-from-xl.xlsx"))

if ($null -ne $error) {
    write-host $error
}
else {
    Convert-CuesToSrt -Cues $cues -OutputPath $([System.IO.Path]::GetFullPath("./samples/[out]/sample-from-xl.xlsx.srt")) | Out-Null
}


# ==============================
## CUES TO EXCEL
Write-Host
Write-Host "CUES TO SRT" -BackgroundColor DarkCyan -ForegroundColor White

$cues = @(
    @{ startMS = 1000; endMS = 5000; text = "Hello, world!" },
    @{ startMS = 6000; endMS = 10000; text = "This is a test." }
)

try {
    Convert-CuesToExcel -Cues $cues -ExcelPath $([System.IO.Path]::GetFullPath("./samples/[out]/output.xlsx")) | Out-Null
}
catch {
    Write-Host "Error Details: $($_.Exception)" -ForegroundColor DarkYellow
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor DarkYellow
}
