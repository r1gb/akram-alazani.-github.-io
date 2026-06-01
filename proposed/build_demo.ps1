<#
  build_demo.ps1
  Usage example:
    powershell -ExecutionPolicy Bypass -File .\build_demo.ps1 -RepoPath "C:\akram-demo\proposed" -OutputDb "C:\akram-demo\database\demo_frontend.accdb"

  This script automates creation of a demo Access frontend on a Windows machine with MS Access installed.
  It imports sample CSV data, imports .bas modules located in proposed/frontend, and attempts to load SaveAsText forms/reports
  if they exist in SaveAsText format. It requires "Trust access to the VBA project object model" enabled in Access
  (File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings).
#>

param(
    [Parameter(Mandatory=$true)] [string] $RepoPath,
    [Parameter(Mandatory=$true)] [string] $OutputDb
)

# Numeric constants for LoadFromText
$acForm = 2
$acReport = 3

# Ensure paths end without backslash
$RepoPath = $RepoPath.TrimEnd('\')
$SampleDataPath = Join-Path $RepoPath "sample_data"
$FrontendPath = Join-Path $RepoPath "frontend"

Try {
    Write-Host "Starting Access COM..."
    $access = New-Object -ComObject Access.Application

    # Create new database (overwrites if exists - be careful)
    if (Test-Path $OutputDb) {
        Write-Host "Output DB exists. Backing up existing file..."
        Copy-Item $OutputDb ($OutputDb + ".bak_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
        Remove-Item $OutputDb
    }

    Write-Host "Creating new database: $OutputDb"
    $access.NewCurrentDatabase($OutputDb)

    # Import CSV files using TransferText
    $csvFiles = @{
        "tblProducts" = (Join-Path $SampleDataPath "products.csv")
        "tblCustomers" = (Join-Path $SampleDataPath "customers.csv")
        "tblInvoices" = (Join-Path $SampleDataPath "invoices.csv")
        "tblInvoiceLines" = (Join-Path $SampleDataPath "invoice_lines.csv")
    }

    foreach ($table in $csvFiles.Keys) {
        $csv = $csvFiles[$table]
        if (Test-Path $csv) {
            Write-Host "Importing $csv into $table ..."
            # acImportDelim = 0
            $access.DoCmd.TransferText(0, "", $table, $csv, $true)
        } else {
            Write-Warning "CSV not found: $csv"
        }
    }

    # Import VBA modules (requires Trust access to VBA project model enabled)
    $basFiles = Get-ChildItem -Path $FrontendPath -Filter *.bas -Recurse -ErrorAction SilentlyContinue
    if ($basFiles.Count -gt 0) {
        Write-Host "Importing .bas modules (requires VBA project access to be trusted)..."
        foreach ($f in $basFiles) {
            Write-Host "  Importing $($f.FullName)"
            $access.VBE.ActiveVBProject.VBComponents.Import($f.FullName)
        }
    } else {
        Write-Host "No .bas files found to import."
    }

    # Load forms/reports from text (only works if files are in SaveAsText format)
    $formTxt = Join-Path $FrontendPath "Form_Dashboard.txt"
    if (Test-Path $formTxt) {
        Write-Host "Loading form from text: $formTxt"
        # Note: second parameter is the object name as it will appear in DB
        $access.LoadFromText($acForm, "frmDashboard", $formTxt)
    } else {
        Write-Host "Form_Dashboard.txt not found or not SaveAsText format (placeholder)."
    }

    $reportTxt = Join-Path $FrontendPath "Report_Invoice.txt"
    if (Test-Path $reportTxt) {
        Write-Host "Loading report from text: $reportTxt"
        $access.LoadFromText($acReport, "rptInvoice", $reportTxt)
    } else {
        Write-Host "Report_Invoice.txt not found or not SaveAsText format (placeholder)."
    }

    # Save & close
    Write-Host "Saving and closing database..."
    $access.CloseCurrentDatabase()
    $access.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) > $null
    Write-Host "Done. Database created at: $OutputDb"
    Write-Host "Next: open the DB in Access, import any remaining placeholders/forms, test, then use File -> Save As -> Make ACCDE to produce ACCDE."
}
Catch {
    Write-Error "Error: $($_.Exception.Message)"
    if ($access -ne $null) {
        try { $access.Quit() } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) > $null
    }
    exit 1
}
