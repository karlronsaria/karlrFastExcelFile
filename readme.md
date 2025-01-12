# FastExcelFile

A set of sysadmin tasks used at a place of employment, made easier with the use of a GUI.

## feature

### Create New Month Book

Create a new Excel workbook based on a month of the year, with a worksheet for each day in the month.

- GUI

  ```dos
  .\Start.bat ~\Downloads\__POOL
  ```

  ![2023_04_20_024708](./res/2023_04_20_024708.png)

- PowerShell

  ```powershell
  .\module\PsTool\Get-Scripts.ps1 | % { . $_ }
  New-MsExcelMonthBook
  ```

- Result

  ![2023_04_20_025210](./res/2023_04_20_025210.png)

  ![2023_04_20_025311](./res/2023_04_20_025311.png)

### Copy From Existing Workbook

Make a copy of an existing workbook ``*.xls*`` and perform a find-and-replace on the name of each worksheet.

- GUI

  ```dos
  .\Start.bat ~\Downloads\__POOL
  ```

  ![2023_04_18_170229](./res/2023_04_18_170229.png)

- PowerShell

  ```powershell
  .\module\PsTool\Get-Scripts.ps1 | % { . $_ }
  $myWorkbook = dir ~\Downloads\__POOL\*.xls* | ? { $_.Name -like "*April*" }
  $myWorkbook | ForEach-MsExcelWorksheet -Do { $_.Name = $_.Name -replace "Apr", "Dec" }
  ```

## howto

Run ``Start.bat``, located in the root folder of this project, from any location. Your working directory can contain the Excel files you want to use, or you can pass in as an argument the directory containing them.

### example

Calling ``Start.bat`` from any location:

```dos
path\to\scripts\Start.bat path\to\excelFiles\
```

Calling ``Start.bat`` from location containing Excel files:

```dos
..\scripts\Start.bat
```

## setup

```dos
git clone --recurse-submodules -j8 https://github.com/karlronsaria/karlrFastExcelFile
```

## maintain

```dos
git pull
git submodule update --init --recursive
```

---
[Progress Page](./doc/todo.md)

