# ----------------------------------------------------------------
#    Check how much free space in the UNC path
#    
#    Replace  $Share value with UNC DocLoad path before running
#    Open up Powershell command line.  Copy paste the below 
# -----------------------------------------------------------------
#
$share=""
#
$nwobj=new-object -comobject WScript.Network
$status=$nwobj.mapnetworkdrive("Z:",$share)
$drive=get-psdrive Z
$gb=(1024 * 1024 * 1024)
$free=($drive.free) /$gb 
$used=($drive.used) /$gb
$total=($free+$used)
$totalrounded=([math]::Round($total))
$freerounded=([math]::Round($free))
$freepercent=($free/$total*100)
$freepercentrounded=([math]::Round($freepercent))   
Write-Output "******* Share $share has  << TOTAL space of $totalrounded GB >> ****"
Write-Output "*******  Share $share has << TOTAL FREE space left $freerounded GB which is $freepercentrounded % >> ****"
$status=$nwobj.removenetworkdrive("Z:")

 