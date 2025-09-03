$ErrorActionPreference = "continue" # "stop"

$modules = @("SRTParser", "ExcelReader")

foreach ($module in $modules) {
    # Ensure modules are not cached during development
    if (Get-Module -Name $module) {
        Remove-Module -Name $module -Force
    }
    Import-Module "./modules/$($module).psm1"
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
    Read-Host "`nDone".
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