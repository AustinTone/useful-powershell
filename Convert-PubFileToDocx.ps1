<#
.SYNOPSIS
	Converts Microsoft Publisher .pub files to Word .docx format via an intermediate PDF.

.DESCRIPTION
	This script automates a two-stage conversion: Publisher files (.pub) are first exported to PDF
	using the Microsoft Office Interop Publisher library, then those PDFs are opened and saved as
	Word documents (.docx) using the Microsoft Office Interop Word library. Intermediate PDFs are
	removed after a successful DOCX conversion. Requires Microsoft Publisher and Microsoft Word to
	be installed and accessible via COM.

.PARAMETER Filter
	Specifies the file filter to select Publisher files for conversion.
	This can be a specific file name (e.g., "document.pub") or a wildcard pattern (e.g., "*.pub").

.PARAMETER Recurse
	If specified, searches for Publisher files recursively in all subdirectories that match the filter. If omitted, only the current directory is searched.

.EXAMPLE
	Convert-PubToDocx.ps1 -Filter "C:\Documents\MyFile.pub"
	Converts the specified Publisher file to a Word document.

.EXAMPLE
	Convert-PubToDocx.ps1 -Filter "*.pub"
	Converts all Publisher files in the current directory to Word documents.

.EXAMPLE
	Convert-PubToDocx.ps1 -Filter "*.pub" -Recurse
	Converts all Publisher files in the current directory and all subdirectories to Word documents.
#>

param
(
    [ValidateNotNullOrEmpty()]
    [string]
    $Filter,

    [switch]
    $Recurse
)

## Functions
function Convert-PubFileToPDF {
    # Checks the filter value to make sure only .pub files are specified
    if (-not ($Filter -like "*.pub")) {
        Write-Error "The filter must specify .pub files (e.g., '*.pub' or 'file.pub').";
        exit 1;
    }

    # Enumerate files specified by -Filter to build an array of target files
    try {
        $files = Get-ChildItem $Filter -File -Recurse:$Recurse;
        if (-not $files) {
            Write-Error "No Publisher files found for the filter: $Filter";
            exit 1;
        }

        # Begin conversion
        Write-Output "Running Publisher to PDF conversion...";

        Add-type -AssemblyName Office;
        Add-type -AssemblyName Microsoft.Office.Interop.Publisher;
        try {
            $app = New-Object -ComObject Publisher.Application;
        }
        catch {
            Write-Error "Microsoft Publisher is not installed or accessible.";
            exit 1;
        }

        $successCount = 0;
        $failCount = 0;

        # First checks if each file has the .pub extension
        foreach ($file in $files) {
            if ($file.Extension -eq ".pub") {
                $fileFullName = $file.FullName;
                $pdfFilePath = [System.IO.Path]::ChangeExtension($fileFullName, '.pdf')
                if (Test-Path $pdfFilePath) {
                    Write-Error "PDF file already exists: $pdfFilePath";
                    $failCount++;
                    Continue;
                }

                # Open the file
                try {
                    $doc = $app.Open($fileFullName);
                }
                catch {
                    $failCount++;
                    Write-Error "Error opening file: $fileFullName $_";
                    Continue;
                }

                if (-not($doc)) {
                    $failCount++;
                    Write-Error "Failed to open file: $fileFullName";
                    Continue;
                }

                try {
                    # Export file as PDF
                    $doc.ExportAsFixedFormat([Microsoft.Office.Interop.Publisher.PbFixedFormatType]::pbFixedFormatTypePDF, $pdfFilePath);
                    if (Test-Path $pdfFilePath) {
                        Write-Output "Exported to $pdfFilePath.";
                        $successCount++;
                    }
                    else {
                        $failCount++;
                        Write-Error "Failed to export file: $fileFullName";
                    }
                }
                catch {
                    $failCount++;
                    Write-Error "Error during export: $_";
                }

                $doc.Close();
            }
        }

        #Log output
        Write-Output "Converted $successCount files with $failCount errors.";
    }
    catch {
        Write-Error $_;
    }
    finally {
        if ($app) {
            #Quit Publisher
            $app.Quit();
        }
    }

}

function Convert-PDFToWord {
    # Enumerate files specified by -Filter to build an array of target files
    $files = Get-ChildItem *.pdf -File -Recurse:$Recurse;
    if (-not $files) {
        Write-Error "No PDF files found.";
        exit 1;
    }

    # Begin conversion
    Write-Output "Running PDF to Word DOCX conversion..."

    Add-type -AssemblyName Office
    Add-Type -AssemblyName Microsoft.Office.Interop.Word

    try {
        $wordApp = New-Object -ComObject Word.Application
        $wordApp.Visible = $false
        $wordApp.DisplayAlerts = 0
    }
    catch {
        Write-Error "Word not installed or accessible";
        exit 1;
    }

    foreach ($file in $files) {
        if ($file.Extension -eq ".pdf") {
            $fileFullName = $file.FullName;
            $docxFilePath = [System.IO.Path]::ChangeExtension($fileFullName, '.docx')
            if (Test-Path $docxFilePath) {
                Write-Error "Word file already exists: $docxFilePath";
                Continue;
            }

            try {
                $pdf = $wordApp.Documents.Open($fileFullName, $false); # False should turn off any conversion confirmation boxes but hasn't been working in testing.
            }
            catch {
                # Write to error stream
                Write-Error "Error opening file: $fileFullName $_";
                Continue;
            }

            if (-not($pdf)) {
                Write-Error "Failed opening file: $fileFullName";
                Continue
            }

            try {
                $pdf.SaveAs2($docxFilePath, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault)
                if (Test-Path $docxFilePath) {
                    Write-Output "Exported to $docxFilePath."
                }
                else {
                    Write-Error "Failed to export file: $fileFullName"
                }
            }
            catch {
                # Write export error
                Write-Error "Error during export: $_"
            }

            $pdf.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
            Remove-Item $fileFullName
        }
    }

    if ($wordApp) {
        # Quit Word
        $wordApp.Quit()
    }
}

## MAIN

# Checks for required -Filter parameter
if (-not $PSBoundParameters.ContainsKey('Filter')) {
    Write-Error "The -Filter parameter is required."
    exit 1
}

# Checks for reg key disabling conversion confirmation boxes created by Word
try {
    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"
    Write-Output "Disabling Word conversion confirmation boxes..."

    if (-not (Get-ItemProperty -Path $regPath -Name "DisableConvertPdfWarning" -ErrorAction SilentlyContinue)) {
        New-ItemProperty -Path $regPath -Name "DisableConvertPdfWarning" -PropertyType DWord -Value 1 | out-null
    }
}
catch {
    # write error
    Write-Error $_
    exit 1
}

# Conversion logic
try {
    Convert-PubFileToPDF

    Convert-PDFToWord
}
catch {
    Write-Error $_
}
finally {
    # Revert set registry key
    if (Get-ItemProperty -Path $regPath -Name "DisableConvertPdfWarning" -ErrorAction SilentlyContinue) {
        Write-Output "Re-enabling Word conversion confirmation boxes..."
        Remove-ItemProperty -Path $regPath -Name "DisableConvertPdfWarning"
    }

    Write-Output "Complete!"
}