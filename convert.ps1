param (
    [string]$folderName
)

if (-not $folderName) {
    Write-Host "Error: Please provide a folder name as an argument."
    exit 1
}

$inputDir = "$PSScriptRoot\$folderName"
$outputDir = "$PSScriptRoot\output"

# Check if input directory exists
if (!(Test-Path $inputDir)) {
    Write-Host "Error: input directory does not exist"
    exit 1
}

# Create output directory if it doesn't exist
if (!(Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir
}

# Check for Excel files
$excelFiles = Get-ChildItem -Path $inputDir -Filter "*.xls"
if ($excelFiles.Count -eq 0) {
    Write-Host "Error: No .xls files found in input directory"
    exit 1
}

# Function to open workbook with retry logic
function Open-WorkbookWithRetry {
    param (
        [string]$filePath,
        [int]$maxRetries = 3,
        [int]$delaySeconds = 5
    )
    $retryCount = 0
    while ($retryCount -lt $maxRetries) {
        try {
            return $excel.Workbooks.Open($filePath)
        }
        catch {
            if ($_.Exception.HResult -eq 0x80010001) {
                Write-Host "Call was rejected by callee. Retrying in $delaySeconds seconds..."
                Start-Sleep -Seconds $delaySeconds
                $retryCount++
            }
            else {
                throw
            }
        }
    }
    throw "Failed to open workbook after $maxRetries retries."
}

# Process Excel files
Write-Host "Processing Excel files..."
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $batchSize = 5
    for ($i = 0; $i -lt $excelFiles.Count; $i += $batchSize) {
        $batch = $excelFiles[$i..[math]::Min($i + $batchSize - 1, $excelFiles.Count - 1)]
        $workbooks = @()
        foreach ($file in $batch) {
            try {
                Write-Host "Opening file $($file.Name)"
                $workbook = Open-WorkbookWithRetry -filePath $file.FullName
                $workbooks += $workbook
            }
            catch {
                Write-Host "Error opening file $($file.Name): $_"
            }
        }
        foreach ($workbook in $workbooks) {
            try {
                $workbook.Save()
                $workbook.Close($true)
                Write-Host "Finished processing file $($workbook.Name)"
            }
            catch {
                Write-Host "Error processing file $($workbook.Name): $_"
            }
        }
    }
}
catch {
    Write-Host "Error: $_"
}
finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    }
}

Write-Host "Done processing files"
