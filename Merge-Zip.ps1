# This is designed to be run either from a folder of scripts, or from any folder of files.
#
# Takes a zip file (InputFile) and adds its compressed files to another zip file (TargetFile)
# preserve relative path info?
# 
# if the file is already exists in the target zip, skip it (and notify to user)
#
# Enjoy!

param ([string]$InputFile,[string]$TargetFile)

# validation
$InputFilePath = Resolve-Path -Path $InputFile

# load .Net Zip assembly
[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression') | Out-Null
[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null

# Handle Target zip file
[System.IO.Compression.ZipArchive] $TargetZipFile = [System.IO.Compression.ZipFile]::Open((Join-Path (Get-Location).ProviderPath $TargetFile),[System.IO.Compression.ZipArchiveMode]::Update)
$TargetFilePath = Resolve-Path -Path $TargetFile

# add your files to archive
[System.IO.Compression.ZipArchive] $InputZipFile = [System.IO.Compression.ZipFile]::OpenRead($InputFilePath.ProviderPath)
$CompressionLevel = [System.IO.Compression.CompressionLevel]::Fastest
$StartTime = Get-Date
$InputZipFile.Entries | 
  ForEach-Object {
    New-Item -ItemType Directory -Force -Path (Split-Path (Join-Path (Get-Location).ProviderPath $_.FullName) -Parent)  | Out-Null
    if($_.Name)  {[System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, (Join-Path (Get-Location).ProviderPath $_.FullName), $false)
      if ($TargetZipFile.Entries.FullName -notcontains $_.FullName) {
        $ZipEntry = $TargetZipFile.CreateEntry($_.FullName,$CompressionLevel)
        $ZipEntryWriter = New-Object -TypeName System.IO.BinaryWriter $ZipEntry.Open()
        $ZipEntryWriter.Write([System.IO.File]::ReadAllBytes((Join-Path (Get-Location).ProviderPath $_.FullName)))
        $ZipEntryWriter.Flush()
        $ZipEntryWriter.Close()
        Write-Host "*** Added $($_.Name) to $(Split-Path $TargetFilePath.ProviderPath -Leaf)" -ForegroundColor Green   
      }
      else {
        Write-Host "!!! File $($_.Name) already exists in Target Zip $(Split-Path $TargetFilePath.ProviderPath -Leaf)" -ForegroundColor Yellow 
      } 
      if ($TargetZipFile.Entries.FullName -like $_.FullName) {Remove-Item (Join-Path (Get-Location).ProviderPath $_.FullName)}
    }
    Get-ChildItem ((Get-Location).ProviderPath) -Recurse | ForEach-Object {
      if( $_.psiscontainer -eq $true){
        if(((Get-ChildItem $_.FullName) -eq $null) -and ($_.LastWriteTime -gt $StartTime)){$_.FullName | Remove-Item -Force}
      }
    }
  }

$TargetZipFile.Dispose()
$InputZipFile.Dispose()
