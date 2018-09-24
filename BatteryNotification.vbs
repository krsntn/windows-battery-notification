set oLocator = CreateObject("WbemScripting.SWbemLocator")
set oServices = oLocator.ConnectServer(".","root\wmi")
set oResults = oServices.ExecQuery("select * from batteryfullchargedcapacity")
for each oResult in oResults
   iFull = oResult.FullChargedCapacity
next

while (1)
  set oResults = oServices.ExecQuery("select * from batterystatus")
  for each oResult in oResults
    iRemaining = oResult.RemainingCapacity
    bCharging = oResult.Charging
  next
  iPercent = ((iRemaining / iFull) * 100)
  if bCharging and (iPercent >= 96) Then msgbox "Battery has charged over 96%"
  if NOT bCharging and (iPercent <= 18) Then msgbox "Battery is getting low"
  wscript.sleep 270000 ' 4.5 minutes
wend