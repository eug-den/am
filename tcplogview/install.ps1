$gser = get-service tcplogview -erroraction silentlycontinue #stop
if ($gser -eq $null)
{
  copy nssm.exe       $env:systemroot
  copy tcplogview.exe $env:systemroot
  copy tcplogview.cfg $env:systemroot
  cmd /c $env:systemroot\nssm.exe install tcplogview $env:systemroot\tcplogview.exe
  start-service -name tcplogview
}
else
{
  "tcplogview service already installed"
}