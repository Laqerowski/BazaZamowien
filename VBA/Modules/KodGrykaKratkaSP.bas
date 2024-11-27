Attribute VB_Name = "KodGrykaKratkaSP"
Public Sub SprawdzamGrykaKratka(ByVal Target As Range, ByVal Arkusz As Worksheet)
    Dim Komorka As Range
    Dim Wiersz As Long

    Application.EnableEvents = False ' Wy³¹cz zdarzenia, aby unikn¹æ zapêtlenia

    ' SprawdŸ ca³¹ kolumnê RODZAJ (B)
    For Wiersz = 2 To Arkusz.Cells(Arkusz.Rows.Count, "B").End(xlUp).Row
        If Arkusz.Cells(Wiersz, "B").Value = "Profilowana" Or Arkusz.Cells(Wiersz, "B").Value = "Gryka sypana" Then
            ' Jeœli w kolumnie RODZAJ jest "Profilowana" lub "Gryka sypana", zmieñ wartoœæ w komórce na "NIE"
            Arkusz.Cells(Wiersz, "C").Value = "NIE" ' Zamieñ wartoœæ w kolumnie GRYKA KRATKA
        End If
    Next Wiersz

    Application.EnableEvents = True ' W³¹cz zdarzenia ponownie
End Sub


