﻿#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $realCmdLet = @{
        OutFile = Get-Command Out-File
    }

    $testInputFile = @{
        Source      = @{
            Folder             = (New-Item 'TestDrive:/s' -ItemType Directory).FullName
            MatchFileNameRegex = 'Analyse_[0-9]{8}.xlsx'
        }
        Destination = @{
            Folder = (New-Item 'TestDrive:/d' -ItemType Directory).FullName
        }
        Settings    = @{
            ScriptName     = 'Test (Brecht)'
            SendMail       = @{
                When         = 'Always'
                From         = 'm@example.com'
                To           = '007@example.com'
                Subject      = 'Email subject'
                Body         = 'Email body'
                Smtp         = @{
                    ServerName     = 'SMTP_SERVER'
                    Port           = 25
                    ConnectionType = 'StartTls'
                    UserName       = 'bob'
                    Password       = 'pass'
                }
                AssemblyPath = @{
                    MailKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                    MimeKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
                }
            }
            SaveLogFiles   = @{
                What                = @{
                    SystemErrors     = $true
                    AllActions       = $true
                    OnlyActionErrors = $false
                }
                Where               = @{
                    Folder         = (New-Item 'TestDrive:/log' -ItemType Directory).FullName
                    FileExtensions = @('.json', '.csv')
                }
                deleteLogsAfterDays = 1
            }
            SaveInEventLog = @{
                Save    = $true
                LogName = 'Scripts'
            }
        }
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ConfigurationJsonFile = $testOutParams.FilePath
    }

    Function Copy-ObjectHC {
        <#
        .SYNOPSIS
            Make a deep copy of an object using JSON serialization.

        .DESCRIPTION
            Uses ConvertTo-Json and ConvertFrom-Json to create an independent
            copy of an object. This method is generally effective for objects
            that can be represented in JSON format.

        .PARAMETER InputObject
            The object to copy.

        .EXAMPLE
            $newArray = Copy-ObjectHC -InputObject $originalArray
        #>
        [CmdletBinding()]
        Param (
            [Parameter(Mandatory)]
            [Object]$InputObject
        )

        $jsonString = $InputObject | ConvertTo-Json -Depth 100

        $deepCopy = $jsonString | ConvertFrom-Json

        return $deepCopy
    }
    function Send-MailKitMessageHC {
        param (
            [parameter(Mandatory)]
            [string]$MailKitAssemblyPath,
            [parameter(Mandatory)]
            [string]$MimeKitAssemblyPath,
            [parameter(Mandatory)]
            [string]$SmtpServerName,
            [parameter(Mandatory)]
            [ValidateSet(25, 465, 587, 2525)]
            [int]$SmtpPort,
            [parameter(Mandatory)]
            [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
            [string]$From,
            [parameter(Mandatory)]
            [string]$Body,
            [parameter(Mandatory)]
            [string]$Subject,
            [string[]]$To,
            [string[]]$Bcc,
            [int]$MaxAttachmentSize = 20MB,
            [ValidateSet(
                'None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable'
            )]
            [string]$SmtpConnectionType = 'None',
            [ValidateSet('Normal', 'Low', 'High')]
            [string]$Priority = 'Normal',
            [string[]]$Attachments,
            [PSCredential]$Credential
        )
    }

    Mock Send-MailKitMessageHC
    Mock New-EventLog
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ConfigurationJsonFile') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'create an error log file when' {
    It 'the log folder cannot be created' {
        Mock Out-File

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Settings.SaveLogFiles.Where.Folder = 'x:\notExistingLocation'

        $testNewInputFile.Source.Folder = (New-Item 'TestDrive:/source' -ItemType Directory).FullName
        $testNewInputFile.Destination.Folder = (New-Item 'TestDrive:/destination' -ItemType Directory).FullName

        New-Item "$($testNewInputFile.Source.Folder)\Analyse_26032025.xlsx" -ItemType File

        & $realCmdLet.OutFile @testOutParams -InputObject (
            $testNewInputFile | ConvertTo-Json -Depth 7
        )

        .$testScript @testParams

        $LASTEXITCODE | Should -Be 1

        Should -Not -Invoke Out-File
    }
    Context 'the ConfigurationJsonFile' {
        It 'is not found' {
            Mock Out-File

            $testNewParams = $testParams.clone()
            $testNewParams.ConfigurationJsonFile = 'nonExisting.json'

            .$testScript @testNewParams

            $LASTEXITCODE | Should -Be 1

            Should -Not -Invoke Out-File
        }
        Context 'property' {
            It '<_> not found' -ForEach @(
                'Folder', 'MatchFileNameRegex'
            ) {
                Mock Out-File

                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Source.$_ = $null

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                $LASTEXITCODE | Should -Be 1

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($LiteralPath -like '* - Errors.json') -and
                    ($InputObject -like "*Property 'Source.$_' not found*")
                }
            }
            It '<_> not found' -ForEach @(
                'Folder'
            ) {
                Mock Out-File

                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Destination.$_ = $null

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($LiteralPath -like '* - Errors.json') -and
                    ($InputObject -like "*Property 'Destination.$_' not found*")
                }
            }
            It 'Folder <_> not found' -ForEach @(
                'Source', 'Destination'
            ) {
                Mock Out-File

                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_.Folder = 'TestDrive:\nonExisting'

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($LiteralPath -like '* - Errors.json') -and
                    ($InputObject -like "*$_.Folder 'TestDrive:\\nonExisting' not found*")
                }
            }
        }
    }
}
Describe 'when the source folder is empty' {
    It 'no error log file is created' {
        Mock Out-File

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Source.Folder = (New-Item 'TestDrive:/empty' -ItemType Directory).FullName

        & $realCmdLet.OutFile @testOutParams -InputObject (
            $testNewInputFile | ConvertTo-Json -Depth 7
        )

        .$testScript @testParams

        Should -Not -Invoke Out-File
    }
}
Describe 'when there is a file in the source folder' {
    It 'the file is moved to the destination folder' {
        $testNewInputFile = Copy-ObjectHC $testInputFile

        $testNewInputFile.Source.Folder = (New-Item 'TestDrive:/source' -ItemType Directory).FullName
        $testNewInputFile.Destination.Folder = (New-Item 'TestDrive:/destination' -ItemType Directory).FullName

        $testFile = New-Item "$($testNewInputFile.Source.Folder)\Analyse_26032025.xlsx" -ItemType File

        & $realCmdLet.OutFile @testOutParams -InputObject (
            $testNewInputFile | ConvertTo-Json -Depth 7
        )

        .$testScript @testParams

        Get-Item "$($testNewInputFile.Destination.Folder)\AnalysesJour_20250326.xlsx" |
        Should -Not -BeNullOrEmpty

        $testFile | Should -Not -Exist
    }
    It 'the file is moved to the destination folder with the correct name' {
        $testNewInputFile = Copy-ObjectHC $testInputFile

        $testNewInputFile.Source.Folder = (New-Item 'TestDrive:/source' -ItemType Directory).FullName
        $testNewInputFile.Destination.Folder = (New-Item 'TestDrive:/destination' -ItemType Directory).FullName

        New-Item "$($testNewInputFile.Source.Folder)\Analyse_26032025.xlsx" -ItemType File

        & $realCmdLet.OutFile @testOutParams -InputObject (
            $testNewInputFile | ConvertTo-Json -Depth 7
        )

        .$testScript @testParams

        Get-Item "$($testNewInputFile.Destination.Folder)\2025\AnalysesJour_20250326.xlsx" |
        Should -Not -BeNullOrEmpty
    } -Skip
}
Describe 'when a file fails to move' {
    BeforeAll {
        Mock Move-Item {
            throw 'Oops'
        }

        $testNewInputFile = Copy-ObjectHC $testInputFile

        $testNewInputFile.Settings.SaveLogFiles.What.AllActions = $false
        $testNewInputFile.Settings.SaveLogFiles.What.OnlyActionErrors = $true

        $testNewInputFile.Source.Folder = (New-Item 'TestDrive:/source' -ItemType Directory).FullName
        $testNewInputFile.Destination.Folder = (New-Item 'TestDrive:/destination' -ItemType Directory).FullName

        $testFile = New-Item "$($testNewInputFile.Source.Folder)\Analyse_26032025.xlsx" -ItemType File

        & $realCmdLet.OutFile @testOutParams -InputObject (
            $testNewInputFile | ConvertTo-Json -Depth 7
        )

        .$testScript @testParams

        $testLogFiles = Get-ChildItem -Path $testInputFile.Settings.SaveLogFiles.Where.Folder -Recurse -File
    }
    It 'error log files are created for each extension' {
        $testLogFiles | Where-Object { $_.Name -like '* - Action errors.json' } |
        Should -Not -BeNullOrEmpty

        $testLogFiles | Where-Object { $_.Name -like '* - Action errors.csv' } |
        Should -Not -BeNullOrEmpty
    }
    It 'Log file content is correct' {
        $testLogFiles | Where-Object {
            $_.Name -like '* - Action errors.json'
        } |
        Get-Content -Raw |
        Should -BeLike  "*Failed to move file '$($testFile.FullName.replace('\', '\\'))'*Oops*"
    }
    It 'an email is sent when SendMail.When is Always' {
        Should -Invoke Send-MailKitMessageHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($Priority -eq 'High') -and
            ($To -eq $testInputFile.Settings.SendMail.To) -and
            ($Subject -eq "1 error, 1 action, $($testInputFile.Settings.SendMail.Subject)") -and
            ($Body -like "*$($testInputFile.Settings.SendMail.Body)*<th>Actions</th>*<td>1</td>*<th>Action errors</th>*<td>1</td>*<th>System errors</th>*<td>0</td>*<p><i>* Check the attachment(s) for details</i></p>*") -and
            ($Attachments -contains $testLogFiles[0].FullName) -and
            ($Attachments -contains $testLogFiles[1].FullName)
        }
    }
}
