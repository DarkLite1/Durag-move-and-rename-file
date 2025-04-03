﻿#Requires -Version 7

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
    $reportData = [System.Collections.Generic.List[PSObject]]::new()

    try {
        $scriptStartTime = Get-Date

        #region Create log folder
        try {
            $logFolderItem = New-Item -Path $LogFolder -ItemType 'Directory' -Force -EA Stop

            $baseLogName = Join-Path -Path $logFolderItem.FullName -ChildPath (
                '{0} - {1}' -f $scriptStartTime.ToString('yyyy_MM_dd_HHmmss_dddd'), $ScriptName
            )

            $logFile = '{0} - Error.txt' -f $baseLogName
        }
        catch {
            Write-Warning "Failed creating the log folder '$LogFolder': $_"

            try {
                "Failed creating the log folder '$LogFolder': $_" |
                Out-File -FilePath "$PSScriptRoot\..\Error.txt"
            }
            catch {
                Write-Warning "Failed creating fallback error file: $_"
            }

            exit 1
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
                    $yearDestinationFolder = Join-Path @params

                    Write-Verbose "Destination folder '$yearDestinationFolder'"

                    $params = @{
                        LiteralPath = $yearDestinationFolder
                        PathType    = 'Container'
                    }
                    if (-not (Test-Path @params)) {
                        $params = @{
                            Path     = $yearDestinationFolder
                            ItemType = 'Directory'
                            Force    = $true
                        }

                        Write-Verbose 'Create destination folder'

                        $null = New-Item @params
                    }
                }
                catch {
                    throw "Failed to create destination folder '$yearDestinationFolder': $_"
                }
                #endregion

                #region Copy file to destination folder
                try {
                    $params = @{
                        LiteralPath = $file.FullName
                        Destination = "$($yearDestinationFolder)\$newFileName"
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
