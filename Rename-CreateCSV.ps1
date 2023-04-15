# This is designed to be run from a folder files you wish to rename.


param ([string]$Path = "./",              # Set Path to examine. May default to local folder.
       [string[]]$Format = @(),           # Set filename elements, e.g. .\Create-Manifest.ps1 -Format "System ID","EEID","Title","Category"
       [string]$FileNameDelimiter = "_",  # Set filename element separator
       [string]$CsvDelimiter = ",",       # Set delimiter in CSV file output
       [switch]$Recursive)
# validation
$DocPath = Resolve-Path -Path $Path

if ((Get-Item $DocPath.ToString()) -isnot [System.IO.DirectoryInfo]) {
  Write-Host "!!! This tool is not designed to be used on individual files such as $($DocPath.ToString().Split('\')[-1])." -ForegroundColor Red
  return 0;
}

# This is how to name columns dynamically based on filename elements
$TokenCount = $Format.Count
if($Recursive) {
  Get-ChildItem -Path $DocPath -File -Recurse |
  ForEach-Object {
    if (($Count = $_.BaseName.Split($FileNameDelimiter).Count) -gt $TokenCount) {$TokenCount = $Count}
  }
}
else {
  Get-ChildItem -Path $DocPath -File |
  ForEach-Object {
    if (($Count = $_.BaseName.Split($FileNameDelimiter).Count) -gt $TokenCount) {$TokenCount = $Count}
  }
}

for ($num = ($Format.Count + 1); $num -le $TokenCount; $num++) {
  $Format += "Element $num"
}

# Build Manifest
$TokenCount = $Format.Count
$DocManifest = Join-Path $DocPath ("Rename-Files.csv")
if (Test-Path $DocManifest) { Remove-Item $DocManifest } # Delete current version

if($Recursive) {
  Get-ChildItem -Path $DocPath -File -Recurse | 
    Where-Object Extension -NotLike "*ps1" | 
    Where-Object BaseName -NotLike "Rename-Files" | # come back to this; may need to add other filters to exclude different manifest files
    Select-Object PSParentPath,FullName,BaseName,Name,Extension | 
    ForEach-Object {
      $FileTokens = [ordered]@{}
      $FileTokens["Current Path"] = $_.PSParentPath.Split(":")[-1]
      $FileTokens["File Name"] = $_.Name
      $FileTokens["New Path"] = $_.PSParentPath.Split(":")[-1]
      $FileTokens["New Name"] = ""
      for ($num = 0; $num -lt $TokenCount; $num++) {
        $FileTokens[$Format[$num]] = $_.BaseName.Split($FileNameDelimiter)[$num]
      }
      $FileTokens["Extension"] = $_.Extension
      [PSCustomObject]$FileTokens | Export-CSV -Path $DocManifest -NoTypeInformation -Encoding UTF8 -Delimiter $CsvDelimiter -Append
    }
}
else {
  Get-ChildItem -Path $DocPath -File | 
    Where-Object Extension -NotLike "*ps1" | 
    Where-Object BaseName -NotLike "Rename-Files" | # come back to this; may need to add other filters to exclude different manifest files
    Select-Object PSParentPath,FullName,BaseName,Name,Extension | 
    ForEach-Object {
      $FileTokens = [ordered]@{}
      $FileTokens["Current Path"] = $_.PSParentPath.Split(":")[-1]
      $FileTokens["File Name"] = $_.Name
      $FileTokens["New Path"] = $_.PSParentPath.Split(":")[-1]
      $FileTokens["New Name"] = ""
      for ($num = 0; $num -lt $TokenCount; $num++) {
        $FileTokens[$Format[$num]] = $_.BaseName.Split($FileNameDelimiter)[$num]
      }
      $FileTokens["Extension"] = $_.Extension
      [PSCustomObject]$FileTokens | Export-CSV -Path $DocManifest -NoTypeInformation -Encoding UTF8 -Delimiter $CsvDelimiter -Append
    }
}

& $DocManifest

