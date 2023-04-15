# This is designed to be run from a folder with files you wish to rename.
# It ingests a .csv file.  Required comma separated, 
# required columns 'Current Path','File Name','New Path','New Name'. Order does not matter
# required content in File Name and New Name columns
# Should NOT overwrite existing files.


param ([string]$Csv = "Rename-Files.csv")
# validation

$FilesCSV = Import-Csv -Path (Resolve-Path -Path $Csv) | Select-Object 'Current Path','File Name','New Path','New Name'

if ($FilesCsv.'File Name'[0]){
  $FilesCsv | ForEach-Object {
    if ($_.'Current Path') {$CurrentFile = Join-Path $_.'Current Path' $_.'File Name'}
    $ErrorActionPreference = "SilentlyContinue"
    if ($currentPath = Resolve-Path $CurrentFile) {
      if (!($_.'New Path')) {$newPath = Split-path $CurrentPath -Parent}
      else {$newPath = $_.'New Path'}
      if (!(Resolve-Path $newPath)) {
        New-Item -ItemType Directory -Force -Path $newPath | out-null
      }
      if (Test-Path -PathType Container $newPath) {
        if ($_.'New Name'){
          $newPath = Join-Path $newPath $_.'New Name'
          if (Test-Path $newPath) {
            Write-Host "!!! This file already exists - ${$newPath.ToString()} does not contain a new file name - confirm CSV headers and content" -ForegroundColor Red
          }
          else {
            Move-Item -Path $currentPath -Destination $newPath | out-null
          }
        }
        else {
          Write-Host "!!! This entry (${$CurrentPath.ToString()}) does not contain a new file name - confirm CSV headers and content" -ForegroundColor Red
        }       
      }
    } 
  }
}
else {
  Write-Host "!!! This file does not have the correct CSV format. Please ensure header row and column `'File Name`' exist." -ForegroundColor Red
  return 0;
}

