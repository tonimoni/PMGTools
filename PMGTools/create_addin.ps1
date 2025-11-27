# PowerShell script to automate the creation of a PowerPoint Add-in (.ppam)

# --- CONFIGURATION ---
$baseName = "ChartDataAddIn"
$pptmFile = "$PSScriptRoot\$baseName.pptm"
$ppamFile = "$PSScriptRoot\$baseName.ppam"
$vbaFile = "$PSScriptRoot\macro.vba"
$ribbonXmlFile = "$PSScriptRoot\ribbon.xml"

# --- SCRIPT BODY ---

Write-Host "Starting the PowerPoint Add-in creation process..."

# Step 1: Create a new PowerPoint instance and presentation
Write-Host "1. Creating PowerPoint presentation..."
$powerpoint = New-Object -ComObject PowerPoint.Application
$presentation = $powerpoint.Presentations.Add()

# Step 2: Inject the VBA code
Write-Host "2. Injecting VBA macro code..."
try {
    $vbaCode = Get-Content -Path $vbaFile -Raw
    $vbaModule = $presentation.VBProject.VBComponents.Add(1) # 1 = vbext_ct_StdModule
    $vbaModule.CodeModule.AddFromString($vbaCode)

    # Add reference to "Microsoft Office XX.0 Object Library"
    # GUID for Office Object Library
    $presentation.VBProject.References.AddFromGuid("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", 2, 8) | Out-Null
}
catch {
    Write-Error "Failed to inject VBA code. Make sure 'Trust access to the VBA project object model' is enabled in PowerPoint's Trust Center settings."
    $powerpoint.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint)
    exit 1
}

# Step 3: Save the presentation as a PPTM file
Write-Host "3. Saving as a macro-enabled presentation (.pptm)..."
# MsoTriState true = -1
$presentation.SaveAs($pptmFile, 25, -1) # 25 = ppSaveAsOpenXMLPresentationMacroEnabled
$presentation.Close()

# Step 4: Inject the RibbonX XML into the package
Write-Host "4. Injecting custom Ribbon XML..."
try {
    Add-Type -AssemblyName System.IO.Packaging

    $pkg = [System.IO.Packaging.Package]::Open($pptmFile, [System.IO.FileMode]::Open)

    # Define the URI for the custom UI part
    $customUiUri = New-Object System.Uri("/customUI/customUI14.xml", "Relative")
    
    # Define the relationship type
    $relType = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"

    # Get the main presentation part and create the relationship
    $presentationPart = $pkg.GetPart((New-Object System.Uri("/ppt/presentation.xml", "Relative")))
    $presentationPart.CreateRelationship($customUiUri, [System.IO.Packaging.TargetMode]::Internal, $relType)

    # Create the custom UI part and write the ribbon XML to it
    $customUiPart = $pkg.CreatePart($customUiUri, "application/xml")
    $ribbonXmlContent = Get-Content $ribbonXmlFile -Raw
    $stream = $customUiPart.GetStream()
    $writer = New-Object System.IO.StreamWriter($stream)
    $writer.Write($ribbonXmlContent)
    $writer.Flush()
    $writer.Close()

    $pkg.Close()
}
catch {
    Write-Error "Failed to inject RibbonX XML. Error was: $($_.Exception.Message)"
    $powerpoint.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint)
    exit 1
}

# Step 5: Re-open the PPTM and save as the final PPAM Add-in
Write-Host "5. Converting to PowerPoint Add-in (.ppam)..."
$presentation = $powerpoint.Presentations.Open($pptmFile, $false, $false, $false) # ReadOnly, Untitled, WithWindow
$presentation.SaveAs($ppamFile, 27) # 27 = ppSaveAsOpenXMLAddin
$presentation.Close()

# Step 6: Cleanup
Write-Host "6. Cleaning up temporary files..."
$powerpoint.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint)
Remove-Item -Path $pptmFile

Write-Host "`nProcess complete!"
Write-Host "Successfully created add-in: $ppamFile"
