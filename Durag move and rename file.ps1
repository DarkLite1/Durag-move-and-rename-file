#Requires -Version 7

<#
    .SYNOPSIS
        Move files from the source folder to the destination folder with a new
        name.

    .DESCRIPTION
        This script selects all files in the source folder that match the
        'MatchFileNameRegex'.

        The selected files are copied from the source folder to the destination
        folder with a new name. The new name is based on the date string
        available withing the source file name.

    .PARAMETER ImportFile
        A .JSON file that contains all the parameters used by the script.

    .PARAMETER SourceFolder
        The source folder.

    .PARAMETER MatchFileNameRegex
        Only files that match the regex will be copied.

    .PARAMETER DestinationFolder
        The destination folder.

    .PARAMETER ProcessFilesInThePastNumberOfDays
        Number of days in the past for which to process files.

        Example:
        - 0 : Process all files in the source folder, no filter
        - 1 : Process files created since yesterday morning
        - 5 : Process files created in the last 5 days

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

            #region Test integer value
            try {
                if ($jsonFileContent.ProcessFilesInThePastNumberOfDays -eq '') {
                    throw 'a blank string is not supported'
                }

                [int]$ProcessFilesInThePastNumberOfDays = $jsonFileContent.ProcessFilesInThePastNumberOfDays

                if ($jsonFileContent.ProcessFilesInThePastNumberOfDays -lt 0) {
                    throw 'a negative number is not supported'
                }
            }
            catch {
                throw "Property 'ProcessFilesInThePastNumberOfDays' must be 0 or a positive number. Number 0 processes all files in the source folder. The value '$($jsonFileContent.ProcessFilesInThePastNumberOfDays)' is not supported."
            }
            #endregion
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
            Recurse     = $true
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
