Attribute VB_Name = "KodGrykaKratkaSP"
Public Sub SprawdzamGrykaKratka(ByVal Target As Range, ByVal Arkusz As Worksheet)
    Dim Komorka As Range
    Dim Wiersz As Long

    Application.EnableEvents = False ' Wy��cz zdarzenia, aby unikn�� zap�tlenia

    ' Sprawd� ca�� kolumn� RODZAJ (B)
    For Wiersz = 2 To Arkusz.Cells(Arkusz.Rows.Count, "B").End(xlUp).Row
        If Arkusz.Cells(Wiersz, "B").Value = "Profilowana" Or Arkusz.Cells(Wiersz, "B").Value = "Gryka sypana" Then
            ' Je�li w kolumnie RODZAJ jest "Profilowana" lub "Gryka sypana", zmie� warto�� w kom�rce na "NIE"
            Arkusz.Cells(Wiersz, "C").Value = "NIE" ' Zamie� warto�� w kolumnie GRYKA KRATKA
        End If
    Next Wiersz

    Application.EnableEvents = True ' W��cz zdarzenia ponownie
End Sub


