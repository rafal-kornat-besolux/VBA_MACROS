Author: Rafał Trachta

Public Sub AddAttachmentNamesToBody()
    Dim xMailItem As Outlook.MailItem
    Dim xAttachment As Attachment
    Dim xInspector As Outlook.Inspector
    Dim xDoc As Word.Document
    Dim xWdSelection As Word.Selection
    Dim my_array(50)
    On Error Resume Next
    Set xMailItem = Outlook.ActiveInspector.CurrentItem
    If xMailItem.Attachments.Count = 0 Then
        Exit Sub
    End If
    l = 0
    For Each xAttachment In xMailItem.Attachments
    
    'my_array(l) = Replace(Mid(Left(xAttachment.FileName, InStrRev(xAttachment.FileName, ".") - 1), 10), "_", "/")
     my_array(l) = Replace(Right(Left(xAttachment.FileName, InStrRev(xAttachment.FileName, ".") - 1), 19), "_", "/")

        l = l + 1
    Next xAttachment
nb = UBound(my_array)
temp_array = my_array
Erase my_array
For i = 0 To nb
    pos = 0
    For l = 0 To nb
        If LCase(temp_array(i)) > LCase(temp_array(l)) And i <> l Then
            pos = pos + 1
        End If
    Next
    For ii = 1 To 1
        If my_array(nb - pos) = "" Then
            my_array(nb - pos) = temp_array(i)
        Else
            pos = pos + 1
            ii = ii - 1
        End If
    Next
Next

If xMailItem.Subject = "" Then

MsgBox "neweail"
Select Case Right(my_array(0), 3)
Case "PKM"
xMailItem.To = "biuro@pkmebel.pl"
Case "ATB"
xMailItem.To = "tomasz@banasiak.pl"
Case "TEC"
xMailItem.To = "abaranowska@tecomat.pl"
Case "AMI"
xMailItem.To = "info@amibelle.pl"
Case "TEX"
xMailItem.To = "i.jasiniak@texpol.net.pl"
Case "MBS"
xMailItem.To = "agnieszka.gzik@mebas.pl"
Case "BEX"
xMailItem.To = "export10@benix.pl; weronika.bielecka@benix.pl"
Case "LEP"
xMailItem.To = "k.badylak@lech-pol.wroc.pl"
Case "NEF"
xMailItem.To = "Export@stolwitmeble.pl"
Case "CPT"
xMailItem.To = "maja@spawmetal.com"
Case "STX"
xMailItem.To = "order@stolmax.info"
Case "MGR"
xMailItem.To = "magdar.meble@wp.pl"
Case "AGP"
xMailItem.To = "tomek.w@agpol.com"
Case "MBP"
xMailItem.To = "biuro@meblepara.pl"
Case "ANG"
xMailItem.To = "dlewkowicz@agnella.pl"
Case "CHX"
xMailItem.To = "chojmex@chojmex.pl"
Case "MLE"
xMailItem.To = "labels@top-line.dk"
Case "DAS"
xMailItem.To = "ewa@domartstyl.pl; magazyn@domartstyl.pl"
Case "ARO"
xMailItem.To = "Casea.pietraszek@arco.net.pl; d.borkowski@arco.net.pl"
Case "DPI"
xMailItem.To = "biuro@dappi.pl"
Case "MIS"
xMailItem.To = "paulina@misofa.pl"
Case "PSM"
xMailItem.To = "logistyka@puszman.com"
Case "MRZ"
xMailItem.To = "export@meble-marzenie.pl; agata.kubelec@meble-marzenie.pl"
Case "DWN"
xMailItem.To = "j.krzciuk@dywilan.pl; zamowienia@dywilan.pl"
Case "RME"
xMailItem.To = "remorse@remorse.pl"
Case "BMR"
xMailItem.To = "ewelina.matusik@bomarmeble.pl"
Case "MAS"
xMailItem.To = "office@maslanka.sale"
Case "NEO"
xMailItem.To = "karolina.cholewa@neo-spiro.pl; j.steclik@neo-spiro.pl"
Case "DAV"
xMailItem.To = "m.majewska@davis.pl"
Case "GIB"
xMailItem.To = "kacperberski.gibmeble@gmail.com"
End Select
xMailItem.CC = "logistics@besolux.com"
xMailItem.Subject = "Etykiety"

End If

    Set xInspector = Outlook.Application.ActiveInspector()
    Set xDoc = xInspector.WordEditor
    Set xWdSelection = xDoc.Application.Selection
    xWdSelection.HomeKey Unit:=wdStory
    
    
    
    For l = 0 To 60
    If my_array(l) <> "" Then
    xWdSelection.InsertBefore my_array(l) & vbCrLf
    End If
    Next l
    
    xWdSelection.InsertAfter vbCrLf & "Prosimy o wyłączenie skalowania w ustawieniach wydruku etykiet, format etykiety zawarty jest w nazwie pliku." & vbCrLf
    xWdSelection.InsertAfter "Jeśli etykieta po wydrukowaniu nie mieści się na Państwa formatce prosimy o informację zwrotną." & vbCrLf & vbCrLf
    
    xWdSelection.InsertAfter "Etykieta oznaczająca produkt:" & vbCrLf
    xWdSelection.InsertAfter "- nie może być porwana" & vbCrLf
    xWdSelection.InsertAfter "- nie może być pomięta" & vbCrLf
    xWdSelection.InsertAfter "- nie może mieć przerw w wydruku" & vbCrLf
    xWdSelection.InsertAfter "- nie może mieć przebarwień" & vbCrLf
    xWdSelection.InsertAfter "- musi być równo naklejona" & vbCrLf
    xWdSelection.InsertAfter "- tekst na etykiecie musi być czytelny" & vbCrLf
    xWdSelection.InsertAfter "- kod kreskowy musi być rozpoznawalny przez skaner"
      
    xWdSelection.InsertBefore xFileName
    xWdSelection.InsertBefore "W załączniku etykiety do zamówień:" & vbCrLf & vbCrLf
    xWdSelection.InsertBefore "Dzień dobry," & vbCrLf & vbCrLf
    
    Set xMailItem = Nothing
    End Sub
