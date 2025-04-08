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
    $eventLogData = [System.Collections.Generic.List[PSObject]]::new()
    $scriptStartTime = Get-Date

    try {
        $eventLogData.Add(
            [PSCustomObject]@{
                DateTime  = $scriptStartTime
                Message   = 'Script started'
                EntryType = 'Information'
                EventID   = '100'
            }
        )

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
            Message  = "Input file '$ImportFile': $_"
        }

        Write-Warning $systemErrors[0].Message

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

        $eventLogData.Add(
            [PSCustomObject]@{
                DateTime  = $scriptStartTime
                Message   = "Found $($filesToProcess.Count) file(s) in source folder '$SourceFolder'"
                EntryType = 'Information'
                EventID   = '4'
            }
        )

        if (!$filesToProcess) {
            Write-Verbose 'No files found, exit script'
            exit
        }
        #endregion

        #region Process files
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

        $eventLogData.Add(
            [PSCustomObject]@{
                DateTime  = $scriptStartTime
                Message   = "Processed $($logFileData.Count) file(s) in source folder '$SourceFolder'"
                EntryType = 'Information'
                EventID   = '4'
            }
        )
        #endregion
    }
    catch {
        $systemErrors += [PSCustomObject]@{
            DateTime = Get-Date
            Message  = $_
        }

        Write-Warning $systemErrors[0].Message

        return
    }
}

end {
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

    function Resolve-LogFolderHC {
        <#
        .SYNOPSIS
            Ensures that a specified path exists, creating it if it doesn't.
            Supports absolute paths and paths relative to $PSScriptRoot.

        .DESCRIPTION
            This function takes a path as input and checks if it exists. If
            the path does not exist, it attempts to create the folder. It handles
            both absolute paths and paths relative to the location of the currently
            running script ($PSScriptRoot).

        .PARAMETER Path
            The path to ensure exists. This can be an absolute path (ex.
            C:\MyFolder\SubFolder) or a path relative to the script's
            directory (ex. Data\Logs).

        .EXAMPLE
            Resolve-LogFolderHC -Path 'C:\MyData\Output'
            # Ensures the directory 'C:\MyData\Output' exists.

        .EXAMPLE
            Resolve-LogFolderHC -Path 'Logs\Archive'
            # If the script is in 'C:\Scripts', this ensures 'C:\Scripts\Logs\Archive' exists.

        .NOTES
            If the path already exists, no action is taken.
            If the creation of the path fails (e.g., due to permissions),
            an error will be thrown.
        #>

        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        if ($Path -match '^[a-zA-Z]:\\' -or $Path -match '^\\') {
            $fullPath = $Path
        }
        else {
            $fullPath = Join-Path -Path $PSScriptRoot -ChildPath $Path
        }

        if (-not (Test-Path -Path $fullPath)) {
            try {
                Write-Verbose "Create log folder '$fullPath'"
                $null = New-Item -Path $fullPath -ItemType Directory -Force
            }
            catch {
                throw "Failed creating log folder '$fullPath': $_"
            }
        }
    }

    function Write-EventsToEventLogHC {
        <#
        .SYNOPSIS
            Write events to the event log.

        .DESCRIPTION
            The use of this function will allow standardization in the Windows
            Event Log by using the same EventID's and other properties across
            different scripts.

            Custom Windows EventID's based on the PowerShell standard streams:

            PowerShell Stream     EventIcon    EventID   EventDescription
            -----------------     ---------    -------   ----------------
            [i] Info              [i] Info     100       Script started
            [4] Verbose           [i] Info     4         Verbose message
            [1] Output/Success    [i] Info     1         Output on success
            [3] Warning           [w] Warning  3         Warning message
            [2] Error             [e] Error    2         Fatal error message
            [i] Info              [i] Info     199       Script ended successfully

        .PARAMETER Source
            Specifies the script name under which the events will be logged.

        .PARAMETER LogName
            Specifies the name of the event log to which the events will be
            written. If the log does not exist, it will be created.

        .PARAMETER Events
            Specifies the events to be written to the event log. This should be
            an array of PSCustomObject with properties: DateTime, Message,
            EntryType, and EventID.
        #>

        [CmdLetBinding()]
        param (
            [Parameter(Mandatory)]
            [String]$Source,
            [Parameter(Mandatory)]
            [String]$LogName,
            [PSCustomObject[]]$Events
        )

        try {
            if (
                -not(
                    ([System.Diagnostics.EventLog]::Exists($LogName)) -and
                    [System.Diagnostics.EventLog]::SourceExists($Source)
                )
            ) {
                Write-Verbose "Create event log '$LogName' and source '$Source'"
                New-EventLog -LogName $LogName -Source $Source -ErrorAction Stop
            }

            foreach ($eventItem in $Events) {
                $params = @{
                    LogName     = $LogName
                    Source      = $Source
                    EntryType   = $eventItem.EntryType
                    EventID     = $eventItem.EventID
                    Message     = '{0}: {1}' -f $eventItem.DateTime, $eventItem.Message
                    ErrorAction = 'Stop'
                }

                Write-Verbose "Write event to log '$LogName' source '$Source' with message '$($params.Message)'"

                Write-EventLog @params
            }
        }
        catch {
            throw "Failed to write to event log '$LogName' with source '$Source': $_"
        }
    }

    try {
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
                Resolve-LogFolderHC -Path $logFolder
                #endregion

                #region Create log folder
                try {
                    Write-Verbose "Create log folder '$LogFolder'"

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
                    Write-Verbose "Result $($logFileData.Count) action(s)"

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
                    Message  = "Failed creating log file in folder '$($jsonFileContent.Settings.Log.Where.Folder)': $_"
                }

                Write-Warning $systemErrors[0].Message
            }
        }

        #region Write events to event log
        if ($logToEventLog) {
            try {
                $systemErrors | ForEach-Object {
                    $eventLogData.Add(
                        [PSCustomObject]@{
                            DateTime  = $_.DateTime
                            Message   = $_.Message
                            EntryType = 'Error'
                            EventID   = '2'
                        }
                    )
                }

                $eventLogData.Add(
                    [PSCustomObject]@{
                        DateTime  = Get-Date
                        Message   = 'Script ended'
                        EntryType = 'Information'
                        EventID   = '199'
                    }
                )

                $params = @{
                    Source  = $scriptName
                    LogName = 'HCScripts'
                    Events  = $eventLogData
                }
                Write-EventsToEventLogHC @params
            }
            catch {
                $systemErrors += [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Failed writing events tot event log $_"
                }

                Write-Warning $systemErrors[0].Message
            }
        }
        else {
            Write-Verbose "Input file option 'Settings.Log.Where.EventLog' not true, no events created in the event log."
        }
        #endregion
    }
    catch {
        $systemErrors += [PSCustomObject]@{
            DateTime = Get-Date
            Message  = $_
        }

        Write-Warning $systemErrors[0].Message
    }
    finally {
        #region Send email

        #endregion

        if ($systemErrors) {
            Write-Warning "Found $($systemErrors.Count) system errors"

            $systemErrors | ForEach-Object {
                Write-Warning $_.Message
            }

            Write-Warning 'Exit script with error code 1'
            exit 1
        }
    }
}
