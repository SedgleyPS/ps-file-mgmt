Get-ChildItem -Path (Get-Location) -Exclude doclist.* | 
  Where-Object { !$_.PSisContainer } |
  Select-Object Name, Length, LastWriteTime | 
  Export-Csv -Path (Join-Path (Get-Location) "doclist.csv") -NoTypeInformation -Encoding UTF8 -Delimiter ","
  Invoke-Item -Path (Join-Path (Get-Location) "doclist.csv")