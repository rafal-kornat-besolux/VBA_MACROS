Public Sub VP_mail()
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.ActiveInspector.CurrentItem
msgbody = OutMail.Body
TO_TEXT = Left(msgbody, InStr(3, msgbody, "Pozdrawiam") - 1)
OutMail.Subject = TO_TEXT
If InStr(1, TO_TEXT, "ADER") <> 0 Then
OutMail.To = "transporte.es@veepee.com"
CC_TEXT = "logistica@aderonline.com ;  mriba@aderonline.com ; asoldevila@aderonline.com ;   alejandro.bielsa@veepee.com ; israel.dura@veepee.com ; veepeebarcelona@aderonline.com"
ElseIf (InStr(1, TO_TEXT, "COLLISIMO") <> 0) Or (InStr(1, TO_TEXT, "VIR")) Or (InStr(1, TO_TEXT, "COLISSIMO") <> 0) OR (InStr(1, TO_TEXT, "COLISIMO") <> 0) Then
OutMail.To = "transport_dropshipment@veepee.com"
CC_TEXT = ""
ElseIf (InStr(1, TO_TEXT, "DHL_DE") <> 0) Or (InStr(1, TO_TEXT, "SPERRGUT") <> 0) Or (InStr(1, TO_TEXT, "DHL DE") <> 0) Or (InStr(1, TO_TEXT, "HOMEDELIVERY") <> 0) Or (InStr(1, TO_TEXT, "HOMEDEL") <> 0) Then
OutMail.To = "transport_dropshipment@veepee.com"
CC_TEXT = "celia.damade@dhl.com ; p.orain@dhl.com ; pauline.mellul@dhl.com"
ElseIf InStr(1, TO_TEXT, "MRW") <> 0 Then
OutMail.To = "transporte.es@veepee.com"
CC_TEXT = "plataforma.barcelona@mrw.es ; operacionescorporatebcn@mrw.es ; Norberto.Sanz@mrw.es ; Victor.Marcos@mrw.es ; alejandro.bielsa@veepee.com ; israel.dura@veepee.com "
ElseIf InStr(1, TO_TEXT, "ASM") <> 0 Then
OutMail.To = "transporte.es@veepee.com"
CC_TEXT = "plataforma.bcn@gls-spain.es ; gema.torres@gls-spain.es ; omar.vives@gls-spain.es ; eva.arribas@gls-spain.es ; israel.dura@veepee.com"
ElseIf InStr(1, TO_TEXT, "POST IT") <> 0 Then
OutMail.To = "transport.italy@veepee.com"
CC_TEXT = "m.lanzon@sda.it ; r.lajolo@sda.it"
ElseIf InStr(1, TO_TEXT, "GLS") <> 0 Then
OutMail.To = "transport.italy@veepee.com"
CC_TEXT = "roberto.tatti@gls-italy.com ; annalisa.varesano@gls-italy.com"
ElseIf InStr(1, TO_TEXT, "LIVOTTI") <> 0 Then
OutMail.To = "transport.italy@veepee.com"
CC_TEXT = "lorenzo.marchegiani@loghistes.it ; magazzino@loghistes.it"
ElseIf InStr(1, TO_TEXT, "NOVATI") <> 0 Then
OutMail.To = "transport.italy@veepee.com"
CC_TEXT = "logistics@besolux.com ; sabrina.pozzi@novatitrasporti.it"
End If

OutMail.CC = CC_TEXT & " ; logistics@besolux.com ; mariaalejandra.perez@veepee.com ; camila.bedoya@veepee.com ; luisa.romero@veepee.com"

text_end = "Pozdrawiam / Kind Regards" & vbCrLf & "__" & vbCrLf
'text_end = text_end & "Konrad Syguła"
'text_end = text_end & "Sebastian Kosiński"
'text_end = text_end & "Marcin Kiełbik"
text_end = text_end & vbCrLf & "www.besolux.com" & vbCrLf & vbCrLf & "Office PL: BESOLUX  //  ul. Łąkowa 7a / bud. E  //  90-562 Łódź  //  Poland" & vbCrLf
text_end = text_end & "Warehouse: BESOLUX  //  ul. Polskich Kolei Państwowych 6  //  92-402 Łódź  //  Poland" & vbCrLf & "NIP: 7292718480"
OutMail.Body = vbCrLf & vbCrLf & text_end
End Sub
