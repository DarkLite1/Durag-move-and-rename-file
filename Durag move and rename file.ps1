#Requires -Version 7

<#
    .SYNOPSIS
        Move files from the source folder to a year folder in the destination folder and rename the file.

    .DESCRIPTION
        This script selects all files in the source folder that match the
        'MatchFileNameRegex'.

        For each selected file the script determines the correct destination
        year destination folder based on the date string in the file name. The
        script also renames the file based on the date string in the file name.

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
        The folder where the log files will be saved.

        Example:
        - Value '..\\Logs'        : Path relative to the script.
        - Value 'C:\\MyApp\\Logs' : An absolute path.
        - Value NULL              : Create no log file.

    .PARAMETER LogFileExtension
        The value is ignored when LogFolder is NULL.

        - Value '.xlsx' : Create an Excel log file.
        - Value '.txt'  : Create a text log file.
        - Value '.csv'  : Create a comma separated log file

    .PARAMETER LogToEventLog
        - Value TRUE : Log verbose to event log.
        - Value FALSE : Do not log messages to the event log.
#>

[CmdLetBinding()]
param (
    [Parameter(Mandatory)]
    [string]$ImportFile
)

begin {
    $ErrorActionPreference = 'stop'

    $systemErrors = @()
    $logFileData = [System.Collections.Generic.List[PSObject]]::new()
    $scriptStartTime = Get-Date

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
        $systemErrors += [PSCustomObject]@{
            DateTime = Get-Date
            Error    = "Input file '$ImportFile': $_"
        }

        Write-Warning $systemErrors[0].Error

        return
    }
}

process {
    if ($systemErrors) { return }

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

                $result = [PSCustomObject]@{
                    DateTime          = Get-Date
                    SourceFolder      = $SourceFolder
                    SourceFileName    = $file.Name
                    NewFileName       = $null
                    DestinationFolder = $null
                    Moved             = $false
                    Error             = $null
                }

                #region Create new file name
                if ($file.Name -notmatch '^\w+_(\d{2})(\d{2})(\d{4})\.\w+$') {
                    throw "Filename '$($file.Name)' does not match expected pattern 'Prefix_ddMMyyyy.ext'."
                }

                $year = $file.Name.Substring(12, 4)
                $month = $file.Name.Substring(10, 2)
                $day = $file.Name.Substring(8, 2)

                $result.NewFileName = "AnalysesJour_$($year)$($month)$($day).xlsx"

                Write-Verbose "New file name '$($result.NewFileName)'"
                #endregion

                #region Create destination folder
                try {
                    $params = @{
                        Path      = $DestinationFolder
                        ChildPath = $year
                    }
                    $result.DestinationFolder = Join-Path @params

                    Write-Verbose "Destination folder '$($result.DestinationFolder)'"

                    $params = @{
                        LiteralPath = $result.DestinationFolder
                        PathType    = 'Container'
                    }
                    if (-not (Test-Path @params)) {
                        $params = @{
                            Path     = $result.DestinationFolder
                            ItemType = 'Directory'
                            Force    = $true
                        }

                        Write-Verbose 'Create destination folder'

                        $null = New-Item @params
                    }
                }
                catch {
                    throw "Failed to create destination folder '$($result.DestinationFolder)': $_"
                }
                #endregion

                #region Move file to destination folder
                try {
                    $params = @{
                        LiteralPath = $file.FullName
                        Destination = "$($result.DestinationFolder)\$($result.NewFileName)"
                        Force       = $true
                    }

                    Write-Verbose "Move file '$($params.LiteralPath)' to '$($params.Destination)'"

                    Move-Item @params
                }
                catch {
                    throw "Failed to move file '$($params.LiteralPath)' to '$($params.Destination)': $_"
                }
                #endregion

                $result.Moved = $true
            }
            catch {
                Write-Warning $_
                $result.Error = $_
            }
            finally {
                $logFileData.Add($result)
            }
        }
    }
    catch {
        Write-Warning "Failure: $_"

        $systemErrors += [PSCustomObject]@{
            DateTime = Get-Date
            Error    = $_
        }

        return
    }
}

end {
    try {
        function Out-LogFileHC {
            param (
                [Parameter(Mandatory)]
                [PSCustomObject[]]$DataToExport,
                [Parameter(Mandatory)]
                [String]$PartialPath,
                [Parameter(Mandatory)]
                [String[]]$FileExtensions
            )

            $allLogFilePaths = @()

            foreach (
                $fileExtension in
                $FileExtensions | Sort-Object -Unique
            ) {
                $logFilePath = $PartialPath -f $fileExtension

                Write-Verbose "Export '$($DataToExport.Count)' objects to '$logFilePath'"

                switch ($fileExtension) {
                    '.txt' {
                        $DataToExport |
                        Out-File -LiteralPath $logFilePath

                        $allLogFilePaths += $logFilePath
                        break
                    }
                    '.csv' {
                        $params = @{
                            LiteralPath       = $logFilePath
                            Delimiter         = ';'
                            NoTypeInformation = $true
                        }
                        $DataToExport | Export-Csv @params

                        $allLogFilePaths += $logFilePath
                        break
                    }
                    '.xlsx' {
                        $excelParams = @{
                            Path          = $logFilePath
                            AutoNameRange = $true
                            AutoSize      = $true
                            FreezeTopRow  = $true
                            WorksheetName = 'Overview'
                            TableName     = 'Overview'
                            Verbose       = $false
                        }
                        $DataToExport | Export-Excel @excelParams

                        $allLogFilePaths += $logFilePath
                        break
                    }
                    default {
                        throw "Log file extension '$_' not supported. Supported values are '.xlsx', '.txt' or '.csv'."
                    }
                }
            }

            $allLogFilePaths
        }

        $scriptName = $jsonFileContent.Settings.ScriptName
        $logFolder = $jsonFileContent.Settings.Log.Where.Folder
        $logFileExtensions = $jsonFileContent.Settings.Log.Where.FileExtensions
        $logToEventLog = $jsonFileContent.Settings.Log.Where.EventLog
        $logSystemErrors = $jsonFileContent.Settings.Log.What.SystemErrors
        $logAllActions = $jsonFileContent.Settings.Log.What.AllActions
        $logOnlyActionErrors = $jsonFileContent.Settings.Log.What.OnlyActionErrors

        #region Get script name
        if (-not $scriptName) {
            Write-Warning "ScriptName not found in import file, using default."
            $scriptName = 'Default script name'
        }
        #endregion

        if ($logFolder -and $logFileExtensions) {
            try {
                #region Get log folder
                try {
                    if (-not [System.IO.Path]::IsPathRooted($logFolder)) {
                        $logFolder = Resolve-Path -Path (
                            Join-Path -Path $PSScriptRoot -ChildPath $logFolder
                        ) -ErrorAction Stop
                    }
                }
                catch {
                    throw "Failed to resolve log folder: $_"
                }
                #endregion

                #region Create log folder
                try {
                    $logFolderItem = New-Item -Path $LogFolder -ItemType 'Directory' -Force -EA Stop

                    $baseLogName = Join-Path -Path $logFolderItem.FullName -ChildPath (
                        '{0} - {1}' -f $scriptStartTime.ToString('yyyy_MM_dd_HHmmss_dddd'), $ScriptName
                    )
                }
                catch {
                    throw "Failed creating log folder '$LogFolder': $_"
                }
                #endregion

                #region Create log file
                $allLogFilePaths = @()

                if ($logFileData) {
                    Write-Warning "$($logFileData.Count) action results"

                    if ($logAllActions) {
                        Write-Verbose 'Export all results'

                        $params = @{
                            DataToExport   = $logFileData
                            PartialPath    = "$baseLogName - Results{0}"
                            FileExtensions = $logFileExtensions
                        }
                        $allLogFilePaths += Out-LogFileHC @params
                    }
                    elseif ($logOnlyActionErrors) {
                        $logFileDataErrors = $logFileData | Where-Object {
                            $_.Error
                        }

                        if ($logFileDataErrors) {
                            Write-Verbose "$($logFileDataErrors.Count) action errors"
                            Write-Verbose 'Export result errors'

                            $params = @{
                                DataToExport   = $logFileDataErrors
                                PartialPath    = "$baseLogName - Errors{0}"
                                FileExtensions = $logFileExtensions
                            }
                            $allLogFilePaths += Out-LogFileHC @params
                        }
                        else {
                            Write-Verbose 'No action errors'
                        }
                    }
                    else {
                        Write-Warning "Log file option 'AllActions', 'OnlyActionErrors' or 'SystemErrors' not found. No log file created."
                    }
                }

                if ($systemErrors) {
                    Write-Warning "$($systemErrors.Count) system errors found"

                    if ($logSystemErrors) {
                        Write-Verbose 'Export system errors'

                        $params = @{
                            DataToExport   = $systemErrors
                            PartialPath    = "$baseLogName - SystemErrors{0}"
                            FileExtensions = $logFileExtensions
                        }
                        $allLogFilePaths += Out-LogFileHC @params
                    }
                    else {
                        Write-Verbose "Input file option 'Settings.Log.SystemErrors' not true. No log file created."
                    }
                }
                #endregion
            }
            catch {
                $systemErrors += [PSCustomObject]@{
                    DateTime = Get-Date
                    Error    = "Failed creating log file in folder '$($jsonFileContent.Settings.Log.Where.Folder)': $_"
                }

                Write-Warning $systemErrors[0].Error
            }
        }

        #region Write events to event log
        if ($logToEventLog) {

        }
        #endregion
    }
    catch {
        Write-Warning "Failure: $_"

        $systemErrors += [PSCustomObject]@{
            DateTime = Get-Date
            Error    = $_
        }
    }
    finally {
        #region Send email

        #endregion

        if ($systemErrors) {
            Write-Warning $systemErrors
            exit 1
        }
    }
}
