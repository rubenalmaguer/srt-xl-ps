# build.ps1

function Join-Scripts {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModulesPath,
        [Parameter(Mandatory = $true)]
        [string]$MainScriptPath,
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    # Create the modules directory if it doesn't exist
    if (-not (Test-Path $ModulesPath)) {
        Write-Error "Modules directory not found: $ModulesPath"
        return
    }

    # Get all .psm1 files from the modules directory
    $moduleFiles = Get-ChildItem -Path $ModulesPath -Filter "*.psm1"
    
    if ($moduleFiles.Count -eq 0) {
        Write-Error "No PowerShell modules (*.psm1) found in: $ModulesPath"
        return
    }

    # Initialize combined content with an empty string
    $combinedModules = ""

    # Process each module
    foreach ($moduleFile in $moduleFiles) {
        $moduleName = $moduleFile.BaseName
        $moduleContent = Get-Content -Path $moduleFile.FullName -Raw

        # Remove all Export-ModuleMember statements, including those with specific function lists
        $moduleContent = $moduleContent -replace 'Export-ModuleMember.*?(?:\r?\n|$)', ''
        # Remove any trailing empty lines that might be left after removing exports
        $moduleContent = $moduleContent -replace '(?m)^\s*\r?\n', ''

        # Add the module content with region markers
        $combinedModules += @"
#Region $moduleName Functions
# $moduleName Functions
$moduleContent

#EndRegion $moduleName Functions

"@
    }

    # Read and process the main script
    $mainScriptContent = Get-Content -Path $MainScriptPath -Raw

    # Remove any module import statements from the main script
    # This pattern will match Import-Module statements for any module
    $mainScriptContent = $mainScriptContent -replace 'Import-Module .*?\.psm1.*?\r?\n', ''

    # Create the PowerShell script content
    $powershellContent = @"
$combinedModules
#Region Main Script
# Main Script Content
$mainScriptContent
#EndRegion Main Script

Write-Host "`nPress any key to exit..."
`$null = `$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
"@

    # Create the complete CMD wrapper content
    $cmdWrapperContent = @"
@@ echo off
@@ title SRT-EXCEL CONVERTER
@@ chcp 65001 >NUL
@@ echo Starting...
@@ REM echo.
@@ setlocal
@@ set PS_WRAPPER_ARGS=%*
@@ set PS_WRAPPER_PATH=%~f0
@@ if defined PS_WRAPPER_ARGS set PS_WRAPPER_ARGS=%PS_WRAPPER_ARGS:"=\"%
@@ PowerShell -NoExit -sta -Command Invoke-Expression `$('`$args=@(^&{`$args} %PS_WRAPPER_ARGS%);'+[String]::Join([Environment]::NewLine,`$((Get-Content '%PS_WRAPPER_PATH%') -notmatch '^^@@^|^^:^|^^cls'))) & endlocal & exit
{
####################### DRAG-N-DROP WRAPPER #######################
#Requires -version 2.0
Set-StrictMode -version 2.0
`$ErrorActionPreference = "stop"
`$script:selfPath = [environment]::GetEnvironmentVariable("PS_WRAPPER_PATH")
`$script:cmdPid = (gwmi win32_process -Filter "processid='`$pid'").parentprocessid
# To access args/dragged files = `$script:args (e.g. `$script:args[1])

######################### EMBEDDED SCRIPT START ###########################

$powershellContent

############################ EMBEDDED SCRIPT END ##########################
}.Invoke(`$args)
"@

    # Ensure the output directory exists
    $outputDir = Split-Path -Parent $OutputPath
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir | Out-Null
    }

    # Save the combined content
    $cmdWrapperContent | Out-File -FilePath $OutputPath -Encoding UTF8

    Write-Host "Successfully combined $($moduleFiles.Count) modules with the main script into: $OutputPath"
}

# Example usage:
$modulesDir = ".\modules"
$mainFile = ".\main.ps1"
$outputFile = ".\dist\Convert-SRT-Excel.cmd"

Join-Scripts -ModulesPath $modulesDir -MainScriptPath $mainFile -OutputPath $outputFile