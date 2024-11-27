Attribute VB_Name = "KodLateksSG"
Public Sub SprawdzLateks(ByVal Arkusz As Worksheet)
    Dim Wiersz As Long

    Application.EnableEvents = False ' Wy³¹cz zdarzenia, aby unikn¹æ zapêtlenia

    ' Iteracja przez wszystkie wiersze z danymi w kolumnie LATEKS (G)
    For Wiersz = 2 To Arkusz.Cells(Arkusz.Rows.Count, "G").End(xlUp).Row
        ' Jeœli w kolumnie LATEKS (G) jest "0/0", wyczyœæ dane w kolumnie G. LATEKS (H)
        If Arkusz.Cells(Wiersz, "G").Value = "0/0" Then
            Arkusz.Cells(Wiersz, "H").ClearContents
        End If
    Next Wiersz

    Application.EnableEvents = True ' W³¹cz zdarzenia ponownie
End Sub

