# Requires -Version 5.0
# Change $Path to a ZIP file that exists on your system!
$Path = "$Home\Desktop\Test.zip"

# Change extension filter to a file extension that exists
# Inside your ZIP file
$Filter = '*.sql'

# Change output path to a folder where you want the extracted
# Files to appear
$OutPath = 'C:\ZIPFiles'

# Ensure the output folder exists
$exists = Test-Path -Path $OutPath
if ($exists -eq $false)
{
  $null = New-Item -Path $OutPath -ItemType Directory -Force
}

# Load ZIP methods
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Open ZIP archive for reading
$zip = [System.IO.Compression.ZipFile]::OpenRead($Path)

# Find all files in ZIP that match the filter (i.e. file extension)
$zip.Entries | 
  Where-Object { $_.FullName -like $Filter } |
  ForEach-Object { 
    # extract the selected items from the ZIP archive
    # and copy them to the out folder
    $FileName = $_.Name
    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, "$OutPath\$FileName", $true)
    }

# Close ZIP file
$zip.Dispose()

# Open out folder
explorer $OutPath