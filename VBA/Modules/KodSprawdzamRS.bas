Attribute VB_Name = "KodSprawdzamRS"
Public Sub Sprawdzam_Rodzaj(ByVal Target As Range, ByVal Arkusz As Worksheet)
    Dim Komorka As Range
    Dim Wiersz As Long

    Application.EnableEvents = False ' Wy³¹cz zdarzenia, aby unikn¹æ zapêtlenia

    ' SprawdŸ ca³¹ kolumnê RODZAJ (B)
    For Wiersz = 2 To Arkusz.Cells(Arkusz.Rows.Count, "B").End(xlUp).Row
        If Arkusz.Cells(Wiersz, "B").Value = "Piankowy" Then
            ' Jeœli w kolumnie RODZAJ jest "Piankowy", wyczyœæ dane w SPRÊ¯YNA i R. SPRÊ¯YNA
            Arkusz.Cells(Wiersz, "D").ClearContents ' Wyczyœæ komórkê w kolumnie SPRÊ¯YNA (D)
            Arkusz.Cells(Wiersz, "E").ClearContents ' Wyczyœæ komórkê w kolumnie R. SPRÊ¯YNA (E)
        End If
    Next Wiersz

    Application.EnableEvents = True ' W³¹cz zdarzenia ponownie
End Sub

