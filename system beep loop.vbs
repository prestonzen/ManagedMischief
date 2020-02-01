do 
Set oWS = WScript.CreateObject("WScript.Shell") 
oWS.Run "%comspec% /c echo " & Chr(07), 0, True 
loop 