 Dim message, sapi
 message=InputBox("What do you want me to say?")
 Set sapi=CreateObject("sapi.spvoice")
 sapi.Speak message