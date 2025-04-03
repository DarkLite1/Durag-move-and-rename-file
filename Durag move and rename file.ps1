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

    $terminatingError = $null
    $logFileData = [System.Collections.Generic.List[PSObject]]::new()

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
            $terminatingError = "Failed creating log folder '$LogFolder': $_"
            return
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
            $terminatingError = "Input file '$ImportFile': $_"
            return
        }
    }
    catch {
        $terminatingError = $_
        return
    }
}

process {
    if ($terminatingError) { return }

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

                $result = @{
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
        $terminatingError = $_
        return
    }
}

end {
    #region Write events to event log

    #endregion

    #region Create log files .csv or .xlsx
    Write-Warning $_
    "Failure:`r`n`r`n- $_" | Out-File -FilePath $logFile -Append

    try {
        "Failed creating the log folder '$LogFolder': $_" |
        Out-File -FilePath "$PSScriptRoot\..\Error.txt"
    }
    catch {
        Write-Warning "Failed creating fallback error file: $_"
    }

    #endregion

    #region Send email

    #endregion

    if ($terminatingError) {
        Write-Warning $terminatingError
        exit 1
    }
}
