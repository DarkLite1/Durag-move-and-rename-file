#Requires -Version 7

<#
    .SYNOPSIS
        Move files from the source folder to the destination folder with a new
        name.

    .DESCRIPTION
        This script selects all files in the source folder that match the
        'MatchFileNameRegex'.

        The selected files are moved from the source folder to the destination
        folder with a new name. The new name is based on the date string
        available withing the source file name.

        Example file:
        - source      : 'Source.Folder\Analyse_26032025.xlsx'
        - destination : 'Destination.Folder\2025\AnalysesJour_20250326.xlsx'

    .PARAMETER ImportFile
        A .JSON file that contains all the parameters used by the script.

    .PARAMETER Source.Folder
        The source folder.

    .PARAMETER Source.MatchFileNameRegex
        Only files that match the regex will be copied.

    .PARAMETER Destination.Folder
        The destination folder.

    .PARAMETER LogFolder
        The folder where the error log files will be saved.
#>

[CmdLetBinding()]
param (
    [Parameter(Mandatory)]
    [string]$ImportFile,
    [string]$ScriptName = 'Process computer actions',
    [string]$LogFolder = "$PSScriptRoot\..\Log"
)

begin {
    $ErrorActionPreference = 'stop'

    try {
        $ScriptStartTime = Get-Date

        #region Create log folder
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -EA Stop
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = '{0} - Error.txt' -f (New-LogFileNameHC @LogParams)
        }
        catch {
            Write-Warning $_

            $params = @{
                FilePath = "$PSScriptRoot\..\Error.txt"
            }
            "Failure:`r`n`r`n- Failed creating the log folder '$LogFolder': $_" | Out-File @params

            exit
        }
        #endregion

        try {
            #region Import .json file
            Write-Verbose "Import .json file '$ImportFile'"

            $jsonFileContent = Get-Content $ImportFile -Raw -Encoding UTF8 |
            ConvertFrom-Json
            #endregion

            $SourceFolder = $jsonFileContent.Source.Folder
            $MatchFileNameRegex = $jsonFileContent.Source.MatchFileNameRegex
            $DestinationFolder = $jsonFileContent.Destination.Folder

            #region Test .json file properties
            @(
                'Folder', 'MatchFileNameRegex'
            ).where(
                { -not $jsonFileContent.Source.$_ }
            ).foreach(
                { throw "Property 'Source.$_' not found" }
            )

            @(
                'Folder'
            ).where(
                { -not $jsonFileContent.Destination.$_ }
            ).foreach(
                { throw "Property 'Destination.$_' not found" }
            )
            #endregion

            #region Test folders exist
            @{
                'Source.Folder'      = $SourceFolder
                'Destination.Folder' = $DestinationFolder
            }.GetEnumerator().ForEach(
                {
                    $key = $_.Key
                    $value = $_.Value

                    if (!(Test-Path -LiteralPath $value -PathType Container)) {
                        throw "$key '$value' not found"
                    }
                }
            )
            #endregion
        }
        catch {
            throw "Input file '$ImportFile': $_"
        }
    }
    catch {
        Write-Warning $_
        "Failure:`r`n`r`n- $_" | Out-File -FilePath $logFile -Append
        exit
    }
}

process {
    try {
        #region Get files from source folder
        Write-Verbose "Get files in source folder '$SourceFolder'"

        $params = @{
            LiteralPath = $SourceFolder
            File        = $true
        }
        $filesToProcess = @(Get-ChildItem @params | Where-Object {
                $_.Name -match $MatchFileNameRegex
            }
        )

        if (!$filesToProcess) {
            Write-Verbose 'No files found, exit script'
            exit
        }
        #endregion

        foreach ($file in $filesToProcess) {
            try {
                Write-Verbose "Processing file '$($file.FullName)'"

                #region Create new file name
                if ($file.Name -notmatch '^\w+_(\d{2})(\d{2})(\d{4})\.\w+$') {
                    throw "Filename '$($file.Name)' does not match expected pattern 'Prefix_ddMMyyyy.ext'."
                }

                $year = $file.Name.Substring(12, 4)
                $month = $file.Name.Substring(10, 2)
                $day = $file.Name.Substring(8, 2)

                $newFileName = "AnalysesJour_$($year)$($month)$($day).xlsx"

                Write-Verbose "New file name '$newFileName'"
                #endregion

                #region Create destination folder
                try {
                    $params = @{
                        Path      = $DestinationFolder
                        ChildPath = $year
                    }
                    $destinationFolder = Join-Path @params

                    Write-Verbose "Destination folder '$destinationFolder'"

                    $params = @{
                        LiteralPath = $destinationFolder
                        PathType    = 'Container'
                    }
                    if (-not (Test-Path @params)) {
                        $params = @{
                            Path     = $destinationFolder
                            ItemType = 'Directory'
                            Force    = $true
                        }

                        Write-Verbose 'Create destination folder'

                        $null = New-Item @params
                    }
                }
                catch {
                    throw "Failed to create destination folder '$destinationFolder': $_"
                }
                #endregion

                #region Copy file to destination folder
                try {
                    $params = @{
                        LiteralPath = $file.FullName
                        Destination = "$($destinationFolder)\$newFileName"
                        Force       = $true
                    }

                    Write-Verbose "Copy file '$($params.LiteralPath)' to '$($params.Destination)'"

                    Copy-Item @params
                }
                catch {
                    throw "Failed to copy file '$($params.LiteralPath)' to '$($params.Destination)': $_"
                }
                #endregion
            }
            catch {
                Write-Warning $_
                "Failure for source file '$($file.FullName)':`r`n`r`n- $_" | Out-File -FilePath $logFile -Append
            }
        }
    }
    catch {
        Write-Warning $_
        "Failure:`r`n`r`n- $_" | Out-File -FilePath $logFile -Append
    }
}

end {

}
