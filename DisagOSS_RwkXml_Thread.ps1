$scriptPath = $PSScriptRoot + "\DisagOSS_RwkXml_Converter.ps1"
$timeout = new-timespan -Hours 24
$sw = [diagnostics.stopwatch]::StartNew()
while ($sw.elapsed -lt $timeout){
    Invoke-Expression $scriptPath
    start-sleep -seconds 2
}
 
write-host "Timed out after:".$timeout