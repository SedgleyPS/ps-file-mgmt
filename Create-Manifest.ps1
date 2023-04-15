# This is designed to be run either from a folder of scripts, or from any folder of files.


param ([string]$Path,                     # Set Path to examine. May default to local folder.
       [string[]]$Format = @(),           # Set filename elements, e.g. .\Create-Manifest.ps1 -Format "System ID","EEID","Title","Category"
       [string]$FileNameDelimiter = "_",  # Set filename element separator
       [string]$CsvDelimiter = ",")       # Set delimiter in CSV file output

# validation
if (!$Path) {$Path = (Read-Host -Prompt "Enter Path (. for current folder)")}
$Path = Resolve-Path -Path $Path

if ((Get-Item $Path.ToString()) -isnot [System.IO.DirectoryInfo]) {
  Write-Host "!!! Create-Manifest.ps1 is not designed to be used on individual files such as $($Path.ToString().Split('\')[-1])." -ForegroundColor Red
  return;
}

# This is how to name columns dynamically based on filename elements
$TokenCount = $Format.Count
Get-ChildItem -Path $Path -File |
  ForEach-Object {
    if (($Count = $_.BaseName.Split($FileNameDelimiter).Count) -gt $TokenCount) {$TokenCount = $Count}
  }
for ($num = ($Format.Count + 1); $num -le $TokenCount; $num++) {
  $Format += "FileName Element $num"
}

# Build Manifest
$TokenCount = $Format.Count
$FileManifest = Join-Path $LogFileDir ("FileInfo" + "_" + (Get-Date -Format "yyyyMMdd") + ".csv")
if (Test-Path $FileManifest) { Remove-Item $FileManifest } # Delete current version

Get-ChildItem -Path $Path -File | 
  Where-Object Extension -NotLike "*ps1" | 
  Where-Object BaseName -NotLike "FileInfo*" | # come back to this; may need to add other filters to exclude different manifest files
  Select-Object BaseName,Extension,CreationTime,Length | 
  ForEach-Object {
    $FileTokens = [ordered]@{}
    for ($num = 0; $num -lt $TokenCount; $num++) {
      $FileTokens[$Format[$num]] = $_.BaseName.Split($FileNameDelimiter)[$num]
    }
    $FileTokens["File Extension"] = $_.Extension
    $FileTokens["Creation Time"] = $_.CreationTime
    $FileTokens["Length"] = $_.Length 
    [PSCustomObject]$FileTokens | Export-CSV -Path $FileManifest -NoTypeInformation -Encoding UTF8 -Delimiter $CsvDelimiter -Append
  } 


