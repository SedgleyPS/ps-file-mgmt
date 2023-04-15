# This is designed to be run either from a folder of scripts, or from any folder of files.
# Easy method for extracts - place this .ps1 in your folder, and run it (right click, run with powershell)
# it will ask for the path, just type a period (.)
# 
# It will ask for an output filename (can include path)
# 
# End result will be a zip containing all files in the directory except for the powershell script file
#
# Enjoy!

param ([string]$Path,[string]$OutputFile)

# validation
if (!$Path) {$Path = (Read-Host -Prompt "Enter Path (. for current folder)")}
$Path = Resolve-Path -Path $Path
if (!$OutputFile) {$OutputFile = (Read-Host -Prompt "Enter the Output Zip Filename")}

# load .Net Zip assembly
[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression') | Out-Null
[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null

# zip files
$OutputZip = Join-Path $Path.ProviderPath $OutputFile
[System.IO.Compression.ZipArchive] $ZipFile = [System.IO.Compression.ZipFile]::Open($OutputZip,[System.IO.Compression.ZipArchiveMode]::Update)

# add your files to archive
Get-ChildItem -Path $Path.ProviderPath -File | 
  Where-Object Extension -NotLike "*ps1" | 
  Where-Object Name -NotLike $OutputZip.Split('\')[-1] |
  ForEach-Object {
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($ZipFile,$_.FullName,$_.Name) | Out-Null
    if ($ZipFile.Entries.FullName -match $_) {
      Write-Host "*** Added $($_.Name) to $($OutputZip.Split('\')[-1])" -ForegroundColor Green
      Remove-Item -LiteralPath $_.FullName
    } 
  }
$ZipFile.Dispose()
