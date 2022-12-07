<# 
.SYNOPSIS
    Extensions to the Microsoft Power BI Management 

.DESCRIPTION 
    Useful functions to help develop Microsoft Power BI reports
 
.NOTES 
    Workspace Id is required for most functions.

.COMPONENT 
    Information about PowerShell Modules to be required.

.LINK 
    Source - https://github.com/7ft10/SevenFt10.PBi.Powershell/blob/main/7ft10PowerBIMgmtUtils.ps1
#>

Function Login-PBICustomPowerBIMgmt
{
    begin 
    {
         if (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt) {
            $checkmodule = "MicrosoftPowerBIMgmt"
            $updateRequired = $false
            $version = (Get-Module -ListAvailable $checkmodule) | Sort-Object Version -Descending  | Select-Object Version -First 1
            $stringver = $version | Select-Object @{n='ModuleVersion'; e={$_.Version -as [string]}}
            $a = $stringver | Select-Object Moduleversion -ExpandProperty Moduleversion

            $psgalleryversion = Find-Module -Name $checkmodule | Sort-Object Version -Descending | Select-Object Version -First 1
            $onlinever = $psgalleryversion | select @{n='OnlineVersion'; e={$_.Version -as [string]}}
            $b = $onlinever | Select-Object OnlineVersion -ExpandProperty OnlineVersion

            $charCount = ($a.ToCharArray() | Where-Object {$_ -eq '.'} | Measure-Object).Count
            switch($charCount)
            {
                {$charCount -eq 1}{ if ([version]('{0}.{1}' -f $a.split('.')) -ge [version]('{0}.{1}' -f $b.split('.'))) { $updateRequired = $false } else { $updateRequired = $true } }
                {$charCount -eq 2}{ if ([version]('{0}.{1}.{2}' -f $a.split('.')) -ge [version]('{0}.{1}.{2}' -f $b.split('.'))) { $updateRequired = $false } else { $updateRequired = $true } }
                {$charCount -eq 3}{ if ([version]('{0}.{1}.{2}.{3}' -f $a.split('.')) -ge [version]('{0}.{1}.{2}.{3}' -f $b.split('.'))) { $updateRequired = $false } else { $updateRequired = $true } }
            }
            if ($updateRequired) {
                Write-Host "Updating..." -NoNewline
                Update-Module -Name MicrosoftPowerBIMgmt > $null
                Write-Host "Done."
            } else {
                Write-Host "PBI Package already installed and updated."
            }
        } 
        else {
            Write-Host "Installing PBI Package..." -NoNewline
            Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser > $null
            Write-Host "Done."
        }
    }

    process 
    {    
        Login-PowerBIServiceAccount > $null
    }

    end {
        Write-Host ""
    }
}

Function Upload-PBIDataSets {
    
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $WorkspaceId,
    
         [Parameter(Mandatory=$true, Position=1)]
         [String] $UploadPathFolderOrFile,

         [Parameter(Mandatory=$true, Position=2)]
         #Ignore - will give the workspace 2 reports with the same name
         #Abort - will error and leave current report uploaded
         #Overwrite - will only upload report if it already exists and overwrites it
         #CreateOrOverwrite - will upload report and overwrite if it does exist
         [Microsoft.PowerBI.Common.Api.Reports.ImportConflictHandlerModeEnum] $ConflictAction
    )
    
    begin 
    {
        $workspace = Get-PowerBIWorkspace -Id $WorkspaceId
        Write-Host "Uploading new data sets to $($workspace.Name)"
    }

    process {
        Upload-PBIFiles $WorkspaceId $UploadPathFolderOrFile $ConflictAction "DataSet.*.pbix"
    }
    
    end {
        Write-Host ""
    }
}

Function Upload-PBIDataSuites {
    
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $WorkspaceId,
    
         [Parameter(Mandatory=$true, Position=1)]
         [String] $UploadPathFolderOrFile,

         [Parameter(Mandatory=$true, Position=2)]
         #Ignore - will give the workspace 2 reports with the same name
         #Abort - will error and leave current report uploaded
         #Overwrite - will only upload report if it already exists and overwrites it
         #CreateOrOverwrite - will upload report and overwrite if it does exist
         [Microsoft.PowerBI.Common.Api.Reports.ImportConflictHandlerModeEnum] $ConflictAction
    )
    
    begin 
    {
        $workspace = Get-PowerBIWorkspace -Id $WorkspaceId
        Write-Host "Uploading new data suites to $($workspace.Name)"
    }

    process {
        Upload-PBIFiles $WorkspaceId $UploadPathFolderOrFile $ConflictAction "DataSuite.*.pbix"
    }
    
    end {
        Write-Host ""
    }
}

Function Upload-PBIReports {
    
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $WorkspaceId,
    
         [Parameter(Mandatory=$true, Position=1)]
         [String] $UploadPathFolderOrFile,

         [Parameter(Mandatory=$true, Position=2)]
         #Ignore - will give the workspace 2 reports with the same name
         #Abort - will error and leave current report uploaded
         #Overwrite - will only upload report if it already exists and overwrites it
         #CreateOrOverwrite - will upload report and overwrite if it does exist
         [Microsoft.PowerBI.Common.Api.Reports.ImportConflictHandlerModeEnum] $ConflictAction
    )
    
    begin 
    {
        $workspace = Get-PowerBIWorkspace -Id $WorkspaceId
        Write-Host "Uploading new reports to $($workspace.Name)"
    }

    process {
        Upload-PBIFiles $WorkspaceId $UploadPathFolderOrFile $ConflictAction "Report.*.pbix"
    }
    
    end {
        Write-Host ""
    }
}

Function Upload-PBIFiles {
    
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $WorkspaceId,
    
         [Parameter(Mandatory=$true, Position=1)]
         [String] $UploadPathFolderOrFile,

         [Parameter(Mandatory=$true, Position=2)]
         [Microsoft.PowerBI.Common.Api.Reports.ImportConflictHandlerModeEnum] $ConflictAction,

         [Parameter(Mandatory=$true, Position=3)]
         [String] $Filter
    )
    
    process {
    
        $isFolder = (Get-Item $UploadPathFolderOrFile) -is [System.IO.DirectoryInfo]
        
        if ($isFolder -eq $true)
        {
            $Reports = Get-ChildItem -Path $UploadPathFolderOrFile -Filter $Filter
        } 
        else 
        {
            $Reports = @( Get-Item -Path $UploadPathFolderOrFile )
        }

        foreach ($r in $Reports) 
        {
            $UploadFile = $r.fullname
            Write-Host "Uploading $($UploadFile)..." -NoNewline
            try 
            {
                $res = New-PowerBIReport -Workspace $Workspace -Path $UploadFile -ConflictAction $ConflictAction -ErrorAction SilentlyContinue -ErrorVariable ProcessError
                If ($ProcessError) 
                {
                    $err = Resolve-PowerBIError -Last
                    if ($err.PowerBIErrorInfo) 
                    {
                        if ($ConflictAction -eq [Microsoft.PowerBI.Common.Api.Reports.ImportConflictHandlerModeEnum]::Abort -And $err.PowerBIErrorInfo -eq "PbixAlreadyImported") 
                        {
                            Write-Host "Skipped." -ForegroundColor Yellow
                        }
                        elseif ($ConflictAction -eq [Microsoft.PowerBI.Common.Api.Reports.ImportConflictHandlerModeEnum]::Overwrite -And $err.PowerBIErrorInfo -eq "DuplicatePackageNotFoundError") 
                        {
                            Write-Host "Report does not exist cannot override." -ForegroundColor Red
                        }                        
                        else 
                        {
                            Write-Host $err.PowerBIErrorInfo -ForegroundColor Red
                        }
                    } 
                    else 
                    {
                        Write-Host $err.Message -ForegroundColor Red
                    }
                } 
                else 
                {                    
                    Write-Host "Done"
                }
            } 
            catch 
            {
                $err = Resolve-PowerBIError -Last
                Write-Host $err.PowerBIErrorInfo
            }       
        }
    }
}

Function Delete-PBIDataSetReports
{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [String]
        $WorkspaceId
    )

    begin {    
        $workspace = Get-PowerBIWorkspace -Id $WorkspaceId 
        Write-Host "Removing data set and data suite reports from $($workspace.Name)"
    }

    process 
    {
        $reports = Get-PowerBIReport -WorkspaceId $WorkspaceId

        $deletedCount = 0

        Foreach ($report in $reports)
        {
            if ($report.Name.StartsWith("DataSet.") -Or $report.Name.StartsWith("DataSuite."))
            {
                $reportName = $report.Name
                Write-Host "Deleting report asscoiated with $reportName..." -NoNewline
                try
                {
                    Remove-PowerBIReport -Id $report.Id -WorkspaceId $WorkspaceId | Out-Null
                    Write-Host "Done"
                    $deletedCount++
                }
                catch
                {
                    $err = Resolve-PowerBIError -Last
                    if ($err.PowerBIErrorInfo) 
                    {
                        Write-Host $err.PowerBIErrorInfo -ForegroundColor Red
                    } 
                    else 
                    {
                        Write-Host $err.Message -ForegroundColor Red
                    }
                }
            }
        }

        if ($deletedCount -eq 0) {
            Write-Host "No DataSet or DataSuite reports to remove."
        }
    }

    end {
        Write-Host ""
    }
}

Function Refresh-PBIAllDataSets
{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [String]
        $WorkspaceId
    )

    begin {    
        $workspace = Get-PowerBIWorkspace -Id $WorkspaceId 
        Write-Host "Refreshing data set reports in $($workspace.Name)"
    }

    process 
    {
        $datasets = Get-PowerBIDataSet -WorkspaceId $WorkspaceId

        $MailFailureNotify = @{"notifyOption"="MailOnFailure"}

        Foreach ($dataset in $datasets)
        {
            Write-Host "Refreshing $($dataset.Name)..." -NoNewline
            if ($dataset.IsRefreshable) 
            {                
                $url = $dataset.WebUrl.Replace("https://app.powerbi.com/", "https://api.powerbi.com/v1.0/myorg/") + "/refreshes"
                $res = Invoke-PowerBIRestMethod -Url $url -Method Post -Body $MailFailureNotify -ErrorAction SilentlyContinue -ErrorVariable RefreshFailed
                if ($RefreshFailed) 
                {
                    $err = Resolve-PowerBIError -Last
                    if ($err.PowerBIErrorInfo) 
                    {
                        Write-Host $err.PowerBIErrorInfo -ForegroundColor Yellow
                    } 
                    else 
                    {
                        try {
                            Write-Host ($err.Message.Replace("One or more errors occurred.", "").Replace("Encountered errors when invoking the command:", "") | ConvertFrom-Json).message.Replace("Invalid dataset refresh request.", "").Trim() -ForegroundColor Red
                        } catch {
                            Write-Host $err.Message -ForegroundColor Red
                        }
                    }
                }
                else 
                {
                    Write-Host "Done"
                }
            } 
            else 
            {
                Write-Host "Skipped - Not Refreshable" -ForegroundColor Yellow
            }
        }        
    }

    end {
        Write-Host ""
    }
}

Function Clear-PBIWorkspace
{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [String]
        $WorkspaceId
    )

    begin {    
        $workspace = Get-PowerBIWorkspace -Id $WorkspaceId 
        Write-Host "Removing all data sets, data suites, reports, everything from $($workspace.Name)"
    }

    process 
    {
        $title = "Clean $($workspace.Name)"
        $question = "Are you sure you want to remove everything from $($workspace.Name)?"
        $choices  = "&Yes", "&No"

        $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
        
        if ($decision -eq 0) 
        {
            $reports = Get-PowerBIReport -WorkspaceId $WorkspaceId

            $deletedCount = 0

            Foreach ($report in $reports)
            {
                $reportName = $report.Name
                Write-Host "Deleting $reportName..." -NoNewline
                try
                {
                    Remove-PowerBIReport -Id $report.Id -WorkspaceId $WorkspaceId | Out-Null
                    Write-Host "Done"
                    $deletedCount++
                }
                catch
                {
                    $err = Resolve-PowerBIError -Last
                    if ($err.PowerBIErrorInfo) 
                    {
                        Write-Host $err.PowerBIErrorInfo -ForegroundColor Red
                    } 
                    else 
                    {
                        Write-Host $err.Message -ForegroundColor Red
                    }
                }
            }

            $datasets = Get-PowerBIDataSet -WorkspaceId $WorkspaceId

            Foreach ($dataset in $datasets)
            {
                $reportName = $dataset.Name
                Write-Host "Deleting $reportName..." -NoNewline

                $url = "https://api.powerbi.com/v1.0/myorg/datasets/$($dataset.Id)"
                $res = Invoke-PowerBIRestMethod -Url $url -Method Delete -ErrorAction SilentlyContinue -ErrorVariable DeleteFailed
                if ($DeleteFailed) 
                {
                    $err = Resolve-PowerBIError -Last
                    if ($err.PowerBIErrorInfo) 
                    {
                        Write-Host $err.PowerBIErrorInfo -ForegroundColor Yellow
                    } 
                    else 
                    {
                        try {
                            Write-Host ($err.Message.Replace("One or more errors occurred.", "").Replace("Encountered errors when invoking the command:", "") | ConvertFrom-Json).message.Replace("Invalid dataset refresh request.", "").Trim() -ForegroundColor Red
                        } catch {
                            Write-Host $err.Message -ForegroundColor Red
                        }
                    }
                }
                else 
                {
                    Write-Host "Done"
                    $deletedCount++
                }
            }

            if ($deletedCount -eq 0) {
                Write-Host "Nothing to remove."
            }
        
        } 
        else 
        {
            Write-Host 'Cancelled'
        }
    }

    end {
        Write-Host ""
    }
}


Function Update-PBIDebugParameters
{
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [String]
        $WorkspaceId,
    
        [Parameter(Mandatory=$true, Position=1)]
        [bool] $IsDebug
    )

    begin {    
        $workspace = Get-PowerBIWorkspace -Id $WorkspaceId 
        Write-Host "Updating dataset parameters in $($workspace.Name) to $IsDebug"
    }

    process 
    {
        $datasets = Get-PowerBIDataSet -WorkspaceId $WorkspaceId

        $content = 'application/json'

        Foreach ($dataset in $datasets)
        {
            Write-Host "Updating parameters for $($dataset.Name)..." -NoNewline
           
            $url = $dataset.WebUrl.Replace("https://app.powerbi.com/", "https://api.powerbi.com/v1.0/myorg/") 

            $body = Invoke-PowerBIRestMethod -Url "$url/parameters" -Method Get -ContentType $content 

            $updated = $false

            if ($body.Contains('IsDebug')) 
            {
                $parameters = ConvertFrom-Json $body

                Foreach ($param in $parameters.value)
                {
                    if ($param.name -eq "IsDebug" -And $param.type -eq "Logical" -And $param.currentValue -ne $IsDebug) 
                    {
                        $updatebody = '{
                          "updateDetails": [
                            {
                              "name": "IsDebug",
                              "newValue": "' + $IsDebug + '"
                            }
                          ]
                        }'
                 
                        $res = Invoke-PowerBIRestMethod -Url "$url/Default.UpdateParameters" -Method Post -Body $updatebody -ContentType $content 
                        Write-Host "Done"
                        $updated = $true
                    }
                }
                 
                if ($updated -eq $false)
                {
                    Write-Host "IsDebug already set" -ForegroundColor Yellow
                    $updated = $true
                }
            } 
            
            if ($updated -eq $false)
            {
                Write-Host "Skipped - no IsDebug Parameter" -ForegroundColor Yellow
            }
        }        
    }

    end {
        Write-Host ""
    }
}
