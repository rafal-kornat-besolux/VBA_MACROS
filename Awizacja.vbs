Public Function sprawdz(N_1)
Select Case N_1
Case "PKM"
RETURN_VALUE = "biuro@pkmebel.pl"
Case "ATB"
RETURN_VALUE = "tomasz@banasiak.pl"
Case "TEC"
RETURN_VALUE = "abaranowska@tecomat.pl"
Case "AMI"
RETURN_VALUE = "info@amibelle.pl"
Case "TEX"
RETURN_VALUE = "i.jasiniak@texpol.net.pl"
Case "MBS"
RETURN_VALUE = "agnieszka.gzik@mebas.pl"
Case "BEX"
RETURN_VALUE = "export10@benix.pl; weronika.bielecka@benix.pl"
Case "LEP"
RETURN_VALUE = "k.badylak@lech-pol.wroc.pl"
Case "NEF"
RETURN_VALUE = "Export@stolwitmeble.pl"
Case "CPT"
RETURN_VALUE = "maja@spawmetal.com"
Case "STX"
RETURN_VALUE = "logistics@stolmax.info"
Case "MGR"
RETURN_VALUE = "magdar.meble@wp.pl"
Case "AGP"
RETURN_VALUE = "tomek.w@agpol.com"
Case "MBP"
RETURN_VALUE = "biuro@meblepara.pl"
Case "ANG"
RETURN_VALUE = "dlewkowicz@agnella.pl  ; rzyskowski@agnella.pl ;  jdzienis@agnella.pl"
Case "CHX"
RETURN_VALUE = "chojmex@chojmex.pl"
Case "MLE"
RETURN_VALUE = "maf@nordkomfort.com ;logistic@nordkomfort.com;ls@nordkomfort.com"
Case "DAS"
RETURN_VALUE = "magazyn@domartstyl.pl"
Case "ARO"
RETURN_VALUE = "Casea.pietraszek@arco.net.pl; d.borkowski@arco.net.pl"
Case "DPI"
RETURN_VALUE = "biuro@dappi.pl"
Case "MIS"
RETURN_VALUE = "paulina@misofa.pl"
Case "PSM"
RETURN_VALUE = "logistyka@puszman.com"
Case "MRZ"
RETURN_VALUE = "agata.kubelec@meble-marzenie.pl"
Case "DWN"
RETURN_VALUE = "j.krzciuk@dywilan.pl; zamowienia@dywilan.pl"
Case "RME"
RETURN_VALUE = "remorse@remorse.pl"
Case "BMR"
RETURN_VALUE = "ewelina.matusik@bomarmeble.pl"
Case "MAS"
RETURN_VALUE = "office@maslanka.sale"
Case "NEO"
RETURN_VALUE = "karolina.cholewa@neo-spiro.pl  ; awizo@neo-spiro.pl"
Case "MOD"
RETURN_VALUE = "magazyn@modalto.pl ; export@modalto.pl ;  msz@modalto.pl  ;rs@modalto.pl  ; mb@modalto.pl"
Case "DAV"
RETURN_VALUE = "m.majewska@davis.pl"
Case "GIB"
RETURN_VALUE = "kacperberski.gibmeble@gmail.com"
Case "SOB"
RETURN_VALUE = "warehouse@besolux.com"
Case "KAL"
RETURN_VALUE = "marta.slowinska@kalmar.pl; patrycja.szalasz@kalmar.pl"      
End Select
sprawdz = RETURN_VALUE
End Function
Public Sub send2()
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.ActiveInspector.CurrentItem
msgbody = OutMail.Body
OutMail.CC = "logistics@besolux.com"
OutMail.Subject = "Awizacja odbioru"
Text = Left(msgbody, InStr(3, msgbody, "Pozdrawiam") - 1)
Dim Result() As String
Result() = Split(Text, vbCrLf)
max_result = UBound(Result())
Dim Factory(40) As String
Factory(1) = "PKM"
Factory(2) = "ATB"
Factory(3) = "TEC"
Factory(4) = "AMI"
Factory(5) = "TEX"
Factory(6) = "MBS"
Factory(7) = "BEX"
Factory(8) = "LEP"
Factory(9) = "NEF"
Factory(10) = "CPT"
Factory(11) = "STX"
Factory(12) = "MGR"
Factory(13) = "AGP"
Factory(14) = "MBP"
Factory(15) = "ANG"
Factory(16) = "CHX"
Factory(17) = "MLE"
Factory(18) = "DAS"
Factory(19) = "ARO"
Factory(20) = "DPI"
Factory(21) = "MIS"
Factory(22) = "PSM"
Factory(23) = "MRZ"
Factory(24) = "DWN"
Factory(25) = "RME"
Factory(26) = "BMR"
Factory(27) = "MAS"
Factory(28) = "NEO"
Factory(29) = "MOD"
Factory(30) = "DAV"
Factory(31) = "GIB"
Factory(32) = "SOB"
Factory(33) = "KAL"   
text_in = ""
For i = 1 To 33
    a = 0
    For j = 0 To max_result
        If InStr(Result(j), Factory(i)) <> 0 Then
            text_in = text_in & Result(j) & vbCrLf
            a = a + 1
         End If
    Next j
    If a <> 0 Then
            text_bcc = text_bcc & sprawdz(Factory(i)) & ";"
            text_in = text_in & vbCrLf
    End If
Next i
OutMail.BCC = text_bcc
text_end = "Pozdrawiam / Kind Regards" & vbCrLf & "__" & vbCrLf
'text_end = text_end & "Konrad Syguła"
'text_end = text_end & "Sebastian Kosiński"
'text_end = text_end & "Marcin Kiełbik"
text_end = text_end & vbCrLf & "www.besolux.com" & vbCrLf & vbCrLf & "Office PL: BESOLUX  //  ul. Łąkowa 7a / bud. E  //  90-562 Łódź  //  Poland" & vbCrLf
text_end = text_end & "Warehouse: BESOLUX  //  ul. Polskich Kolei Państwowych 6  //  92-402 Łódź  //  Poland" & vbCrLf & "NIP: 7292718480"
OutMail.Body = "Dzień dobry," & vbCrLf & vbCrLf & "W dniu" & vbCrLf & vbCrLf & "Po zamówienia:" & vbCrLf & vbCrLf & text_in & "Stawi się:" & vbCrLf & vbCrLf & vbCrLf & text_end
End Sub
