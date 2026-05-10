do
Set wshshell = wscript.CreateObject("WScript.Shell")
Wshshell.run "Notepad"
wscript.sleep 100
wshshell.sendkeys "P"
wscript.sleep 100
wshshell.sendkeys "I"
wscript.sleep 100
wshshell.sendkeys "N"
wscript.sleep 100
wshshell.sendkeys "G"
wscript.sleep 100
wshshell.sendkeys "S"
loop