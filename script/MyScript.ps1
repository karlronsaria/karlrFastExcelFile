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
            Name = "NewMonthBook"
            Text = "New Month Book"
        },
        [PsCustomObject]@{
            Name = "CopyExisting"
            Text = "Copy Existing Workbook"
        }
    )
}

$prefs = Get-QformPreference `
    -Preference ([PsCustomObject]@{
        Caption = "Choose What Do"
    })

$months = @{}

$newMonthBook = @(
    [PsCustomObject]@{
        Name = "Year"
        Type = "Numeric"
        Minimum = "1970"
        Maximum = "9999"
        Default = (Get-Date).Year
        Mandatory = $true
    },
    [PsCustomObject]@{
        Name = "Month"
        Type = "Enum"
        Symbols = 1 .. 12 | foreach {
            $name = (Get-Date -Month $_).ToString("MMMM")
            $months[$name] = $_

            [PsCustomObject]@{
                Name = $name
            }
        }
        Default = (Get-Date).ToString("MMMM")
        Mandatory = $true
    }
)

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
    'NewMonthBook' {
        $prefs = Get-QformPreference `
            -Reference $prefs `
            -Preference ([PsCustomObject]@{
                Caption = "New Month Book"
            })

        $what = $newMonthBook | Show-QformMenu `
            -Preferences $prefs

        if (-not $what.Confirm) {
            return
        }

        $answers = $what.MenuAnswers

        $setting = cat "$PsScriptRoot/../res/setting.json" `
            | ConvertFrom-Json

        New-MsExcelMonthBook `
            -Year $answers.Year `
            -Month $months[$answers.Month] `
            -ColumnHeadings $setting.ColumnHeadings
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

