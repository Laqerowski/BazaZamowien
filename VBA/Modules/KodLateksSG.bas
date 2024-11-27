Attribute VB_Name = "KodLateksSG"
Public Sub SprawdzLateks(ByVal Arkusz As Worksheet)
    Dim Wiersz As Long

    Application.EnableEvents = False ' Wy��cz zdarzenia, aby unikn�� zap�tlenia

    ' Iteracja przez wszystkie wiersze z danymi w kolumnie LATEKS (G)
    For Wiersz = 2 To Arkusz.Cells(Arkusz.Rows.Count, "G").End(xlUp).Row
        ' Je�li w kolumnie LATEKS (G) jest "0/0", wyczy�� dane w kolumnie G. LATEKS (H)
        If Arkusz.Cells(Wiersz, "G").Value = "0/0" Then
            Arkusz.Cells(Wiersz, "H").ClearContents
        End If
    Next Wiersz

    Application.EnableEvents = True ' W��cz zdarzenia ponownie
End Sub

