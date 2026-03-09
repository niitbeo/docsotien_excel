# ================================================
# AUTO INSTALL DOCTIEN TO EXCEL
# One-click installation for Windows 11 + Excel
# ================================================

Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  INSTALLING DOCTIEN TO EXCEL..." -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""

# Check if DocTien.bas exists
$basFile = Join-Path $PSScriptRoot "DocTien.bas"

if (-not (Test-Path $basFile)) {
    Write-Host "[ERROR] Cannot find DocTien.bas" -ForegroundColor Red
    Write-Host "Please make sure DocTien.bas is in the same folder." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit
}

Write-Host "[OK] Found DocTien.bas" -ForegroundColor Green
Write-Host ""

# Start Excel
Write-Host "[PROGRESS] Starting Excel..." -ForegroundColor Yellow

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    Write-Host "[OK] Excel ready" -ForegroundColor Green
    Write-Host ""
    
    # Get XLSTART path
    $xlStartPath = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
    
    if (-not (Test-Path $xlStartPath)) {
        New-Item -ItemType Directory -Path $xlStartPath -Force | Out-Null
    }
    
    $personalFile = Join-Path $xlStartPath "PERSONAL.XLSB"
    
    Write-Host "[PROGRESS] Creating Personal Macro Workbook..." -ForegroundColor Yellow
    
    # Open or create PERSONAL.XLSB
    if (Test-Path $personalFile) {
        Write-Host "   -> Opening existing PERSONAL.XLSB" -ForegroundColor Cyan
        $workbook = $excel.Workbooks.Open($personalFile)
    }
    else {
        Write-Host "   -> Creating new PERSONAL.XLSB" -ForegroundColor Cyan
        $workbook = $excel.Workbooks.Add()
    }
    
    Write-Host "[OK] Workbook ready" -ForegroundColor Green
    Write-Host ""
    
    # Remove existing DocTien module if exists
    foreach ($component in $workbook.VBProject.VBComponents) {
        if ($component.Name -eq "DocTienModule") {
            Write-Host "   -> Removing old DocTien module..." -ForegroundColor Yellow
            $workbook.VBProject.VBComponents.Remove($component)
            break
        }
    }
    
    # Import DocTien module
    Write-Host "[PROGRESS] Importing DocTien module..." -ForegroundColor Yellow
    
    try {
        $workbook.VBProject.VBComponents.Import($basFile) | Out-Null
        Write-Host "[OK] Successfully imported DocTien!" -ForegroundColor Green
    }
    catch {
        Write-Host "[ERROR] Failed to import module!" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        Write-Host ""
        Write-Host "[WARNING] You may need to enable VBA project access:" -ForegroundColor Yellow
        Write-Host "   1. Open Excel" -ForegroundColor White
        Write-Host "   2. File -> Options -> Trust Center -> Trust Center Settings" -ForegroundColor White
        Write-Host "   3. Macro Settings -> Check 'Trust access to the VBA project object model'" -ForegroundColor White
        Write-Host "   4. Click OK and run this script again" -ForegroundColor White
        
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        Read-Host "Press Enter to exit"
        exit
    }
    
    Write-Host ""
    
    # Save file
    Write-Host "[PROGRESS] Saving file..." -ForegroundColor Yellow
    
    if (Test-Path $personalFile) {
        $workbook.Save()
    }
    else {
        $workbook.SaveAs($personalFile, 50)  # 50 = xlExcel12 (.xlsb)
    }
    
    Write-Host "[OK] Saved PERSONAL.XLSB" -ForegroundColor Green
    Write-Host ""
    
    # Close Excel
    Write-Host "[PROGRESS] Finishing..." -ForegroundColor Yellow
    $workbook.Close($false)
    $excel.Quit()
    
    # Cleanup
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host "  INSTALLATION COMPLETE!" -ForegroundColor Green
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "You can now use DocTien() function in ANY Excel file!" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "HOW TO USE:" -ForegroundColor Yellow
    Write-Host "   1. Open Excel (or restart if already open)" -ForegroundColor White
    Write-Host "   2. Type formula: =DocTien(A1)" -ForegroundColor White
    Write-Host ""
    Write-Host "EXAMPLE:" -ForegroundColor Yellow
    Write-Host "   =DocTien(12345)" -ForegroundColor White
    Write-Host "   -> Muoi hai nghin ba tram bon muoi lam dong" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "NOTE:" -ForegroundColor Yellow
    Write-Host "   - Click 'Enable Content' if you see a security warning" -ForegroundColor White
    Write-Host "   - File installed at: $personalFile" -ForegroundColor White
    Write-Host ""
    
}
catch {
    Write-Host ""
    Write-Host "[ERROR] Unexpected error:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    
    # Cleanup
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

Write-Host ""
Read-Host "Press Enter to close"
