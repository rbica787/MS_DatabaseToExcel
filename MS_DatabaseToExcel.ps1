# =====================================================================
# Export Access MDB Tables to a Single Excel Workbook
#
# What this script does:
#   - Automatically relaunches itself in 32-bit PowerShell if needed
#   - Prompts for the Microsoft Access .MDB file
#   - Finds all user tables
#   - Exports every table
#   - Creates one Excel worksheet per Access table
#   - Optionally converts every worksheet into an Excel Table
#   - Uses the first row as the Excel Table header
#   - Defaults the Excel Table option to YES if Enter is pressed
#   - Saves everything into one .XLSX file
#   - Uses the MDB filename plus a timestamp
# =====================================================================


# ---------------------------------------------------------------------
# AUTO-RELAUNCH IN 32-BIT POWERSHELL
# ---------------------------------------------------------------------

if ([Environment]::Is64BitProcess) {

    $PowerShell32 = "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"

    if (-not (Test-Path -LiteralPath $PowerShell32)) {

        Write-Host ""
        Write-Host "ERROR: 32-bit PowerShell could not be found." -ForegroundColor Red
        Write-Host ""
        exit
    }

    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {

        Write-Host ""
        Write-Host "ERROR: This script must be saved as a .ps1 file." -ForegroundColor Red
        Write-Host ""
        exit
    }

    Write-Host ""
    Write-Host "64-bit PowerShell detected."
    Write-Host "Restarting script in 32-bit PowerShell..."
    Write-Host ""

    & $PowerShell32 -ExecutionPolicy Bypass -File $PSCommandPath

    exit
}


# ---------------------------------------------------------------------
# CONFIRM 32-BIT POWERSHELL
# ---------------------------------------------------------------------

Write-Host ""
Write-Host "============================================================"
Write-Host " Access MDB -> Excel Export"
Write-Host "============================================================"
Write-Host ""

Write-Host "PowerShell mode:"
Write-Host "  32-bit"
Write-Host ""


# ---------------------------------------------------------------------
# PROMPT FOR MDB FILE
# ---------------------------------------------------------------------

$MdbPath = Read-Host "Enter the full path to the .mdb file"

$MdbPath = $MdbPath.Trim().Trim('"').Trim("'")


# ---------------------------------------------------------------------
# VERIFY FILE EXISTS
# ---------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $MdbPath -PathType Leaf)) {

    Write-Host ""
    Write-Host "ERROR: MDB file was not found:" -ForegroundColor Red
    Write-Host ""
    Write-Host "  $MdbPath" -ForegroundColor Yellow
    Write-Host ""
    exit
}


# ---------------------------------------------------------------------
# VERIFY FILE EXTENSION
# ---------------------------------------------------------------------

if ([System.IO.Path]::GetExtension($MdbPath) -ne ".mdb") {

    Write-Host ""
    Write-Host "ERROR: The selected file is not an .mdb file:" -ForegroundColor Red
    Write-Host ""
    Write-Host "  $MdbPath" -ForegroundColor Yellow
    Write-Host ""
    exit
}


# ---------------------------------------------------------------------
# PROMPT FOR EXCEL TABLES
# ---------------------------------------------------------------------

Write-Host ""
Write-Host "Excel Table option:"
Write-Host ""

do {

    $TableChoice = Read-Host "Would you like the contents of every worksheet converted to an Excel table? (Y/N) [Default: Y]"

    $TableChoice = $TableChoice.Trim().ToUpper()

    # If nothing is entered, default to YES.
    if ([string]::IsNullOrWhiteSpace($TableChoice)) {
        $TableChoice = "Y"
    }

    if ($TableChoice -ne "Y" -and $TableChoice -ne "N") {

        Write-Host ""
        Write-Host "Please enter Y or N. Press Enter for Yes." -ForegroundColor Yellow
        Write-Host ""
    }

}
while ($TableChoice -ne "Y" -and $TableChoice -ne "N")


$CreateExcelTables = ($TableChoice -eq "Y")


Write-Host ""

if ($CreateExcelTables) {

    Write-Host "Excel Tables:"
    Write-Host "  Enabled"
    Write-Host "  Row 1 will be used as the header."
    Write-Host ""

}
else {

    Write-Host "Excel Tables:"
    Write-Host "  Disabled"
    Write-Host ""

}


# ---------------------------------------------------------------------
# BUILD OUTPUT FILE NAME
# ---------------------------------------------------------------------

$MdbPath = (Resolve-Path -LiteralPath $MdbPath).Path

$Folder = Split-Path -Parent $MdbPath
$BaseName = [System.IO.Path]::GetFileNameWithoutExtension($MdbPath)

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

$ExcelPath = Join-Path $Folder "$($BaseName)_$($Timestamp).xlsx"


# ---------------------------------------------------------------------
# OBJECT VARIABLES
# ---------------------------------------------------------------------

$Connection = $null
$Excel = $null
$Workbook = $null


try {

    Write-Host ""
    Write-Host "MDB file:"
    Write-Host "  $MdbPath"
    Write-Host ""


    # -----------------------------------------------------------------
    # CREATE DATABASE CONNECTION
    # -----------------------------------------------------------------

    $Connection = New-Object -ComObject ADODB.Connection

    $Connected = $false


    # -----------------------------------------------------------------
    # TRY ACE 16.0
    # -----------------------------------------------------------------

    try {

        $ConnectionString =
            "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=$MdbPath;Persist Security Info=False;"

        $Connection.Open($ConnectionString)

        $Connected = $true

        Write-Host "Database provider:"
        Write-Host "  Microsoft.ACE.OLEDB.16.0"
        Write-Host ""

    }
    catch {

        Write-Host "ACE 16.0 connection failed."
        Write-Host "Trying ACE 12.0..."
        Write-Host ""
    }


    # -----------------------------------------------------------------
    # TRY ACE 12.0
    # -----------------------------------------------------------------

    if (-not $Connected) {

        try {

            if ($Connection.State -ne 0) {
                $Connection.Close()
            }

        }
        catch {
        }

        try {

            $ConnectionString =
                "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$MdbPath;Persist Security Info=False;"

            $Connection.Open($ConnectionString)

            $Connected = $true

            Write-Host "Database provider:"
            Write-Host "  Microsoft.ACE.OLEDB.12.0"
            Write-Host ""

        }
        catch {

            Write-Host "ACE 12.0 connection failed."
            Write-Host "Trying Microsoft Jet..."
            Write-Host ""
        }
    }


    # -----------------------------------------------------------------
    # TRY JET 4.0
    # -----------------------------------------------------------------

    if (-not $Connected) {

        try {

            if ($Connection.State -ne 0) {
                $Connection.Close()
            }

        }
        catch {
        }

        try {

            $ConnectionString =
                "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$MdbPath;"

            $Connection.Open($ConnectionString)

            $Connected = $true

            Write-Host "Database provider:"
            Write-Host "  Microsoft.Jet.OLEDB.4.0"
            Write-Host ""

        }
        catch {

            throw @"

Unable to open the MDB database.

The script tried:

    Microsoft.ACE.OLEDB.16.0
    Microsoft.ACE.OLEDB.12.0
    Microsoft.Jet.OLEDB.4.0

Database:
$MdbPath

"@
        }
    }


    # -----------------------------------------------------------------
    # FIND ALL TABLES
    # -----------------------------------------------------------------

    Write-Host "Reading table list..."

    # adSchemaTables = 20
    $Schema = $Connection.OpenSchema(20)

    $TableNames = New-Object System.Collections.Generic.List[string]


    while (-not $Schema.EOF) {

        $TableName = [string]$Schema.Fields.Item("TABLE_NAME").Value
        $TableType = [string]$Schema.Fields.Item("TABLE_TYPE").Value

        if (
            $TableType -eq "TABLE" -and
            $TableName -notlike "MSys*"
        ) {

            if (-not $TableNames.Contains($TableName)) {
                $TableNames.Add($TableName)
            }
        }

        $Schema.MoveNext()
    }


    $Schema.Close()

    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
        $Schema
    )


    if ($TableNames.Count -eq 0) {
        throw "No user tables were found in the MDB database."
    }


    Write-Host ""
    Write-Host "Tables found: $($TableNames.Count)"
    Write-Host ""

    foreach ($TableName in $TableNames) {
        Write-Host "  $TableName"
    }

    Write-Host ""


    # -----------------------------------------------------------------
    # START EXCEL
    # -----------------------------------------------------------------

    Write-Host "Starting Excel..."
    Write-Host ""

    $Excel = New-Object -ComObject Excel.Application

    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false

    $Workbook = $Excel.Workbooks.Add()


    # -----------------------------------------------------------------
    # REMOVE EXTRA DEFAULT WORKSHEETS
    # -----------------------------------------------------------------

    while ($Workbook.Worksheets.Count -gt 1) {

        $SheetToDelete = $Workbook.Worksheets.Item(
            $Workbook.Worksheets.Count
        )

        $SheetToDelete.Delete()

        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $SheetToDelete
        )
    }


    # -----------------------------------------------------------------
    # KEEP TRACK OF WORKSHEET NAMES
    # -----------------------------------------------------------------

    $UsedSheetNames = @{}

    $FirstSheet = $true
    $TableNumber = 0


    # -----------------------------------------------------------------
    # EXPORT EACH TABLE
    # -----------------------------------------------------------------

    foreach ($TableName in $TableNames) {

        $TableNumber++

        Write-Host "------------------------------------------------------------"
        Write-Host "[$TableNumber/$($TableNames.Count)] Exporting:"
        Write-Host "  $TableName"


        # -------------------------------------------------------------
        # CREATE SAFE EXCEL SHEET NAME
        # -------------------------------------------------------------

        $SheetName = $TableName

        # Excel sheet names cannot contain:
        # : \ / ? * [ ]
        $SheetName = $SheetName -replace '[:\\\/\?\*\[\]]', '_'


        if ($SheetName.Length -gt 31) {
            $SheetName = $SheetName.Substring(0,31)
        }


        if ([string]::IsNullOrWhiteSpace($SheetName)) {
            $SheetName = "Table"
        }


        # -------------------------------------------------------------
        # MAKE SHEET NAME UNIQUE
        # -------------------------------------------------------------

        $OriginalSheetName = $SheetName
        $Counter = 1


        while ($UsedSheetNames.ContainsKey($SheetName.ToLower())) {

            $Counter++

            $Suffix = "_$Counter"

            $MaximumBaseLength = 31 - $Suffix.Length

            $TempName = $OriginalSheetName


            if ($TempName.Length -gt $MaximumBaseLength) {

                $TempName = $TempName.Substring(
                    0,
                    $MaximumBaseLength
                )
            }

            $SheetName = "$TempName$Suffix"
        }


        $UsedSheetNames[$SheetName.ToLower()] = $true


        # -------------------------------------------------------------
        # CREATE OR REUSE WORKSHEET
        # -------------------------------------------------------------

        if ($FirstSheet) {

            $Worksheet = $Workbook.Worksheets.Item(1)

            $FirstSheet = $false

        }
        else {

            # Create a new worksheet.
            #
            # IMPORTANT:
            # Do not pass Before/After arguments and do not call Move().
            # Both caused COM errors in this Excel installation.

            $Worksheet = $Workbook.Worksheets.Add()
        }


        $Worksheet.Name = $SheetName


        # -------------------------------------------------------------
        # OPEN ACCESS TABLE
        # -------------------------------------------------------------

        $Recordset = New-Object -ComObject ADODB.Recordset

        $EscapedTableName = $TableName.Replace("]", "]]")

        $Sql = "SELECT * FROM [$EscapedTableName]"


        try {

            $Recordset.Open(
                $Sql,
                $Connection,
                0,
                1
            )

        }
        catch {

            Write-Host "  ERROR reading table." -ForegroundColor Red
            Write-Host "  $($_.Exception.Message)" -ForegroundColor Red

            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $Recordset
            )

            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $Worksheet
            )

            continue
        }


        # -------------------------------------------------------------
        # WRITE COLUMN HEADERS
        # -------------------------------------------------------------

        $FieldCount = $Recordset.Fields.Count


        for ($Column = 0; $Column -lt $FieldCount; $Column++) {

            $Field = $Recordset.Fields.Item($Column)

            $Worksheet.Cells.Item(
                1,
                $Column + 1
            ).Value2 = $Field.Name

            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $Field
            )
        }


        # -------------------------------------------------------------
        # FORMAT HEADER ROW
        # -------------------------------------------------------------

        if ($FieldCount -gt 0) {

            $HeaderRange = $Worksheet.Range(
                $Worksheet.Cells.Item(1,1),
                $Worksheet.Cells.Item(1,$FieldCount)
            )

            $HeaderRange.Font.Bold = $true

            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $HeaderRange
            )
        }


        # -------------------------------------------------------------
        # COPY TABLE DATA
        # -------------------------------------------------------------

        $RowsCopied = 0

        if (-not $Recordset.EOF) {

            $StartCell = $Worksheet.Cells.Item(2,1)

            $RowsCopied = $StartCell.CopyFromRecordset(
                $Recordset
            )

            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $StartCell
            )

            Write-Host "  Records exported: $RowsCopied"

        }
        else {

            Write-Host "  Records exported: 0"
        }


        # -------------------------------------------------------------
        # CONVERT WORKSHEET CONTENTS TO EXCEL TABLE
        # -------------------------------------------------------------

        if ($CreateExcelTables -and $FieldCount -gt 0) {

            try {

                # Row 1 contains the field names and becomes the
                # Excel Table header row.
                #
                # The final row is:
                #   Header row + number of exported records

                $LastRow = [int]$RowsCopied + 1


                # Excel requires a table to contain at least one data row.
                # If the Access table contains no records, create one
                # temporary blank row. Excel will keep the table structure.

                $TemporaryBlankRow = $false

                if ($LastRow -lt 2) {

                    $LastRow = 2
                    $TemporaryBlankRow = $true
                }


                $FirstTableCell = $Worksheet.Cells.Item(1,1)

                $LastTableCell = $Worksheet.Cells.Item(
                    $LastRow,
                    $FieldCount
                )

                $TableRange = $Worksheet.Range(
                    $FirstTableCell,
                    $LastTableCell
                )


                # xlSrcRange = 1
                # xlYes      = 1
                #
                # ListObjects.Add(SourceType, Source, LinkSource,
                #                 XlListObjectHasHeaders)

                $ExcelTable = $Worksheet.ListObjects.Add(
                    1,
                    $TableRange,
                    $null,
                    1
                )


                # -----------------------------------------------------
                # CREATE A SAFE, UNIQUE EXCEL TABLE NAME
                # -----------------------------------------------------

                $SafeTableName = $SheetName -replace '[^A-Za-z0-9_]', '_'

                if ([string]::IsNullOrWhiteSpace($SafeTableName)) {
                    $SafeTableName = "Table"
                }


                # Prefixing the name prevents problems when an Access
                # table name resembles an Excel cell address.

                $SafeTableName = "tbl_$SafeTableName"


                # Include the table number to guarantee uniqueness
                # throughout the workbook.

                $SafeTableName = "$($SafeTableName)_$TableNumber"


                # Excel table names are limited to 255 characters.

                if ($SafeTableName.Length -gt 255) {
                    $SafeTableName = $SafeTableName.Substring(0,255)
                }


                $ExcelTable.Name = $SafeTableName


                # If the source table had zero records, remove the blank
                # Excel table data row after the ListObject is created.

                if ($TemporaryBlankRow) {

                    try {

                        if ($ExcelTable.ListRows.Count -gt 0) {
                            $ExcelTable.ListRows.Item(1).Delete()
                        }

                    }
                    catch {

                        # If Excel will not remove the only blank row,
                        # leave it blank. The header/table structure is
                        # still valid.
                    }
                }


                Write-Host "  Excel table: $SafeTableName"


                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $ExcelTable
                )

                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $TableRange
                )

                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $LastTableCell
                )

                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $FirstTableCell
                )

            }
            catch {

                Write-Host "  WARNING: Could not create Excel table." -ForegroundColor Yellow
                Write-Host "  $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }


        # -------------------------------------------------------------
        # FREEZE TOP ROW
        # -------------------------------------------------------------

        try {

            $Worksheet.Activate()

            $Excel.ActiveWindow.SplitRow = 1
            $Excel.ActiveWindow.FreezePanes = $true

        }
        catch {

            # Freeze panes is cosmetic.
            # Continue if Excel refuses it.
        }


        # -------------------------------------------------------------
        # AUTO-FIT COLUMNS
        # -------------------------------------------------------------

        try {

            $UsedRange = $Worksheet.UsedRange

            $UsedRange.Columns.AutoFit() | Out-Null

            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $UsedRange
            )

        }
        catch {

            # AutoFit is cosmetic.
            # Continue if Excel refuses it.
        }


        # -------------------------------------------------------------
        # CLOSE RECORDSET
        # -------------------------------------------------------------

        if ($Recordset.State -ne 0) {
            $Recordset.Close()
        }


        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $Recordset
        )

        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
            $Worksheet
        )


        Write-Host "  Worksheet: $SheetName"
        Write-Host "  Complete."
    }


    # -----------------------------------------------------------------
    # SAVE XLSX
    # -----------------------------------------------------------------

    Write-Host ""
    Write-Host "============================================================"
    Write-Host "Saving Excel workbook..."
    Write-Host "============================================================"
    Write-Host ""


    # xlOpenXMLWorkbook = 51
    $Workbook.SaveAs(
        $ExcelPath,
        51
    )


    Write-Host "EXPORT COMPLETE" -ForegroundColor Green
    Write-Host ""
    Write-Host "Excel file created:"
    Write-Host ""
    Write-Host "  $ExcelPath" -ForegroundColor Cyan
    Write-Host ""

}
catch {

    Write-Host ""
    Write-Host "============================================================"
    Write-Host "ERROR" -ForegroundColor Red
    Write-Host "============================================================"
    Write-Host ""

    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
}
finally {

    # -----------------------------------------------------------------
    # CLOSE DATABASE CONNECTION
    # -----------------------------------------------------------------

    if ($null -ne $Connection) {

        try {

            if ($Connection.State -ne 0) {
                $Connection.Close()
            }

        }
        catch {
        }

        try {

            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $Connection
            )

        }
        catch {
        }
    }


    # -----------------------------------------------------------------
    # CLOSE EXCEL
    # -----------------------------------------------------------------

    if ($null -ne $Workbook) {

        try {
            $Workbook.Close($false)
        }
        catch {
        }

        try {

            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $Workbook
            )

        }
        catch {
        }
    }


    if ($null -ne $Excel) {

        try {
            $Excel.Quit()
        }
        catch {
        }

        try {

            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $Excel
            )

        }
        catch {
        }
    }


    # -----------------------------------------------------------------
    # COM CLEANUP
    # -----------------------------------------------------------------

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}