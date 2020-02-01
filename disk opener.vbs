Set oWMP = CreateObject("WMPlayer.OCX.7")
Set colCDROMs = oWMP.cdromCollection
if colCDROMs.Count = 1 then
for i = 0 to colCDROMS.Count - 1
colCDROMs.Item(i).Eject
colCDROMs.Item(i).Eject
Next ' cdrom