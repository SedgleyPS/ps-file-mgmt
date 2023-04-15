# This is designed to be run either from a folder of scripts, or from any folder of files.


param ([string]$Path,                     # Set Path to examine. 
       [string[]]$Format = @(),           # Set filename elements, e.g. .\Create-Manifest.ps1 -Format "System ID","EEID","Title","Category"
       [string]$FileNameDelimiter = "_",  # Set filename element separator
       [string]$CsvDelimiter = ",")       # Set delimiter in CSV file output

# validation
if (!$Path) {$Path = (Read-Host -Prompt "Enter Zip Path")}
$ZipPath = (Get-ItemProperty -Path (Resolve-Path -Path $Path).Path)

# if the path is not a zip file, don't do anything more!
if (!(Test-Path $ZipPath -PathType Leaf) -or !($ZipPath.Extension -eq ".zip")) {
  Write-Host "!!! \$($ZipPath.Name) is not a zip file - please use a different script." -ForegroundColor Yellow
  return 0;
}

#Build the Manifest file; store alongside zip
$ZipManifest = Join-Path $ZipPath.Directory ("FileInfo" + "_" + $ZipPath.BaseName + ".csv")

#Initialize catalog array
$zipCatalog = @()

function Get-UncompressedZipEntries {
  param ($Path)
  
  $shell = New-Object -ComObject shell.application
  $zip = $shell.NameSpace($Path)
  foreach ($item in $zip.items()) {
    if ($item.IsFolder) {
      Get-UncompressedZipEntries -Path $item.Path
    } 
    else {   
      $zipCatalogEntry = [ordered]@{}
      $zipCatalogEntry['Name'] = $item.Name
      $zipCatalogEntry['Extension'] = "." + $item.Name.Split(".")[-1]
      $zipCatalogEntry['BaseName'] = $item.Name -replace ($zipCatalogEntry['Extension'] + '$'),''
      $zipCatalogEntry['Path'] = ($item.Path.Remove(0,$ZipPath.ToString().Length))
      $zipCatalogEntry['Size'] = $item.Size
      $zipCatalogEntry['Date'] = $item.ModifyDate
      $zipCatalog += , $zipCatalogEntry
    }
  }

  # dispose the COM object now explicitly
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$shell) | Out-Null
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  return $zipCatalog
  
}

#Build Catalog Object
$zipCatalog = Get-UncompressedZipEntries($ZipPath.ToString())

# This is how to name columns dynamically based on filename elements
$TokenCount = $Format.Count
$zipCatalog |
  ForEach-Object {
    if (($Count = $_.BaseName.Split($FileNameDelimiter).Count) -gt $TokenCount) {$TokenCount = $Count}
  }
for ($num = ($Format.Count + 1); $num -le $TokenCount; $num++) {
  $Format += "FileName Element $num"
}
$TokenCount = $Format.Count

# Build Manifest
if (Test-Path $ZipManifest) { Remove-Item $ZipManifest } # Delete current version

$zipCatalog | 
  ForEach-Object {
    if ($_.BaseName -notmatch "^FileInfo" ) {
      $FileTokens = [ordered]@{}
      for ($num = 0; $num -lt $TokenCount; $num++) {
        $FileTokens[$Format[$num]] = $_.BaseName.Split($FileNameDelimiter)[$num]
      }
      $FileTokens["File Path"] = $_.Path -replace ([regex]::Escape($_.Name) + "$"),''
      $FileTokens["File Name"] = $_.BaseName
      $FileTokens["Extension"] = $_.Extension
      $FileTokens["Creation Time"] = $_.Date
      $FileTokens["Length"] = $_.Size 
      [PSCustomObject]$FileTokens | Export-CSV -Path $ZipManifest -NoTypeInformation -Encoding UTF8 -Delimiter $CsvDelimiter -Append
    }
  } 


