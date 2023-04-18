Param(
    $Path
)

if (-not $Path) {
    $Path = (Get-Location).FullName
}

$startIn = "$PsScriptRoot/.."

Import-Module "$startIn/module/PsQuickform/PsQuickform.psm1"
& "$startIn/module/PsTool/Get-Scripts.ps1" | % { . $_ }

$start = [PsCustomObject]@{
    Name = "Start"
    Type = "Enum"
    Text = "What do?"
    # Mandatory = $true
    Symbols = @(
        [PsCustomObject]@{
            Name = "CopyExisting"
            Text = "Copy Existing Workbook"
        },
        [PsCustomObject]@{
            Name = "NewWorkbook"
            Text = "New Workbook"
        }
    )
}

$prefs = Get-QformPreference `
    -Preference ([PsCustomObject]@{
        Caption = "Choose What Do"
    })

$copyExisting = @(
    [PsCustomObject]@{
        Name = "WorkbooksTable"
        Type = "Table"
        Text = "Workbooks in $((Get-Item $Path).FullName)"
        Rows = dir "$Path/*.xls*" -Recurse `
            | select Name, Directory, LastWriteTime
    },
    [PsCustomObject]@{
        Name = "Find"
        Type = "Field"
        Text = "For each worksheet, replace"
    },
    [PsCustomObject]@{
        Name = "Replace"
        Type = "Field"
        Text = "in the name, with"
    }
)

$what = $start | Show-QformMenu `
    -Preferences $prefs

if (-not $what.Confirm) {
    return
}

switch ($what.MenuAnswers.Start) {
    'NewWorkbook' {
        $prefs = Get-QformPreference `
            -Preference ([PsCustomObject]@{
                Caption = "Not Implemented"
            })

        Show-QformMenu -Preferences $prefs
    }

    'CopyExisting' {
        $prefs = Get-QformPreference `
            -Reference $prefs `
            -Preferences ([PsCustomObject]@{
                Caption = "Copy Existing Excel File"
            })

        $what = $copyExisting | Show-QformMenu `
            -Preferences $prefs

        if (-not $what.Confirm) {
            return
        }

        $answers = $what.MenuAnswers
        $workbooks = $answers.WorkbooksTable

        foreach ($workbook in $workbooks) {
            $path = Join-Path $workbook.Directory $workbook.Name

            if (-not $answers.Find) {
                return Get-Item $path
            }

            dir $path | ForEach-MsExcelWorksheet -Do {
                $result = $_.Name -replace $answers.Find, $answers.Replace
                Write-Output "Renaming: `"$($_.Name)`" to `"$result`""
                $_.Name = $result
            }
        }
    }
}

