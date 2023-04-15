param ($Path, [switch]$space)

function Get-UncompressedZipFileSize {
  param ($Path)

  
  $shell = New-Object -ComObject shell.application
  $zip = $shell.NameSpace($Path)
  $size = 0
  foreach ($item in $zip.items()) {
    if ($item.IsFolder) {
      $size += Get-UncompressedZipFileSize -Path $item.Path
    } 
    else {
      $size += $item.size
    }
  }

  # It might be a good idea to dispose the COM object now explicitly, see comments below
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$shell) | Out-Null
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  return $size
  
}

Function Format-FileSize() {
  Param ([int64]$size)
  If     ($size -gt 1TB) {[string]::Format("{0:0.00} TB", $size / 1TB)}
  ElseIf ($size -gt 1GB) {[string]::Format("{0:0.00} GB", $size / 1GB)}
  ElseIf ($size -gt 1MB) {[string]::Format("{0:0.00} MB", $size / 1MB)}
  ElseIf ($size -gt 1KB) {[string]::Format("{0:0.00} kB", $size / 1KB)}
  ElseIf ($size -gt 0)   {[string]::Format("{0:0.00} B", $size)}
  Else                   {""}
}

if ($space) { return Format-FileSize(Get-UncompressedZipFileSize $Path) }

[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null
$zip = [IO.Compression.ZipFile]::OpenRead($Path)
$zip.Entries.FullName | %{ "$Path`:$_" }
$zip.Dispose()

