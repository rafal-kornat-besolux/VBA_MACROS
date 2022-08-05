Public Sub SEND2()
Dim Factory_email                   'Create a variable
Set Factory_email = CreateObject("Scripting.Dictionary")
Factory_email("PKM") = "biuro@pkmebel.pl"
Factory_email.Add "MBS", "agnieszka.gzik@mebas.pl"
Factory_email.Add "BEX", "export10@benix.pl; weronika.bielecka@benix.pl"
Factory_email.Add "LEP", "a.patyk@liic.eu; a.lenart@liic.eu; k.badylak@liic.eu; r.pieprzka@lech-pol.wroc.pl"
Factory_email.Add "CPT", "maja@spawmetal.com; info@spawmetal.com"
Factory_email.Add "SPW", "maja@spawmetal.com; info@spawmetal.com"
Factory_email.Add "STX", "logistics@stolmax.info"
Factory_email.Add "AGP", "tomek.w@agpol.com; tomek.g@agpol.com"
Factory_email.Add "ANG", "dlewkowicz@agnella.pl; jdzienis@agnella.pl; jmiejluk@agnella.pl; rzyskowski@agnella.pl"
Factory_email.Add "CHX", "chojmex@chojmex.pl"
Factory_email.Add "MLE", "maf@nordkomfort.com ;logistic@nordkomfort.com;ls@nordkomfort.com; ordernk@nordkomfort.com"
Factory_email.Add "DAS", "magazyn@domartstyl.pl"
Factory_email.Add "PSM", "export@puszman.com; logistyka@puszman.com"
Factory_email.Add "DPI", "biuro@dappi.pl"
Factory_email.Add "MRZ", "dluczka@comforteo.eu; biuro@comforteo.eu"
Factory_email.Add "BMR", "ewelina.matusik@bomarmeble.pl"
Factory_email.Add "NEO", "sara.nabrdalik@neo-spiro.pl; karolina.nocunfolfasinska@neo-spiro.pl; awizo@neo-spiro.pl"
Factory_email.Add "MOD", "magazyn@modalto.pl ; export@modalto.pl ;  msz@modalto.pl  ;rs@modalto.pl  ; mb@modalto.pl"
Factory_email.Add "DAV", "m.majewska@DAVIS.pl; m.waczynski@davis.pl; a.borkala@DAVIS.pl"
Factory_email.Add "GIB", "eksport@gibmeble.pl; kacperberski.gibmeble@gmail.com"
Factory_email.Add "SOB", "warehouse@besolux.com"
Factory_email.Add "BSO", "warehouse@besolux.com"
Factory_email.Add "CNY", "warehouse@besolux.com"
Factory_email.Add "KAL", "gabriela.jablonska@kalmar.pl; patrycja.szalasz@kalmar.pl"
Factory_email.Add "CHB", "biuro@chemeb.pl"
Factory_email.Add "DOL", "m.dobrzynska@dolmar.pl"
Factory_email.Add "GAL", "estera.wroblewska@galameble.com; grot@galameble.com; agnieszka.wojciechowska@galameble.com"
Factory_email.Add "KAC", "biuro@mt-kaczmarek.com; l.kaczmarek@tessa-sofa.com"
Factory_email.Add "WJW", "biuro@ccloft.pl"
Factory_email.Add "WOJ", "biuro@ccloft.pl"
Factory_email.Add "ZAM", "commerciale@zamagna.com; lzamagna@zamagna.com; fashinterconsulting@gmail.com; tsantonocito@zamagna.com"
Factory_email.Add "PRO", "office@prospero.net.pl"
Factory_email.Add "DES", "office@despro.net.pl"
Factory_email.Add "DRE", "drewmix@home.pl; m.ogrodnik@drewmix.home.pl"
Factory_email.Add "MLS", "ordernk@nordkomfort.com; maf@nordkomfort.com; ls@nordkomfort.com; logistic@nordkomfort.com"
Factory_email.Add "RCO", "steclik.jakub@gmail.com"
Factory_email.Add "STL", "monika.solarz@stolar.com.pl"


Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.ActiveInspector.CurrentItem
msgbody = OutMail.Body
Text = Left(msgbody, InStr(3, msgbody, "Pozdrawiam") - 1)
Dim Result() As String
Result() = Split(Text, vbCrLf)
max_result = UBound(Result())
text_bcc = ""
text_in = ""
text2 = ""
For j = 0 To max_result
    X1 = 0
    For Each varKey In Factory_email.Keys()
        If InStr(Result(j), varKey) <> 0 Then
        X1 = 1
        End If
    Next
    If X1 = 0 And (Len(Result(j)) > 3) Then
    text2 = text2 & "nie ma maila dla" & Result(j) & "<br>"
    End If
Next j

If text2 = "" Then

    For Each varKey In Factory_email.Keys()
        a = 0
        For j = 0 To max_result
            If InStr(Result(j), varKey) <> 0 Then
                text_in = text_in & Result(j) & "<br>"
                a = a + 1
            End If
        Next j
        If a <> 0 Then

                text_bcc = text_bcc + Factory_email(varKey) + ";"
                text_in = text_in & "<br>"
        End If
    Next

    Text = "Dzień dobry," & "<br>" & "<br>" & "W dniu" & "<br>" & "<br>" & "Po zamówienia:" & "<br>" & "<br>" & text_in & "Stawi się:" & "<br>" & "<br>" & "<br>"

    OutMail.Close olSave
    Call X("", "logistics@beslox.com", text_bcc, "Awizacja odbioru", Text + "")
Else
    OutMail.Close olSave
    Call X("", "logistics@beslox.com", "", "Awizacja odbioru", text2 + "")
End If
End Sub

Sub X(strTo As String, strCC As String, strBCC, strSubject As String, strBody As String)

   Dim OlApp As Outlook.Application
   Dim ObjMail As Outlook.MailItem

   Set OlApp = Outlook.Application
   Set ObjMail = OlApp.CreateItem(olMailItem)

   ObjMail.To = strTo
   ObjMail.CC = strCC
   ObjMail.BCC = strBCC
   ObjMail.Subject = strSubject
   ObjMail.Display
   'You now have the default signature within ObjMail.HTMLBody.
   'Add this after adding strHTMLBody
   ObjMail.HTMLBody = strBody + ObjMail.HTMLBody

   'ObjMail.Send 'send immediately or
   'ObjMail.close olSave 'save as draft
   'Set OlApp = Nothing

End Sub
