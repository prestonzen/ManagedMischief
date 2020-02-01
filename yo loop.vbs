set shellobj = CreateObject("WScript.Shell")
shellobj.run "cmd"

do

shellobj.sendkeys "Y"
wscript.sleep 200
Shellobj.sendkeys "o "
wscript.sleep 200

loop