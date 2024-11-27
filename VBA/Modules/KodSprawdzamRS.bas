Attribute VB_Name = "KodSprawdzamRS"
Public Sub Sprawdzam_Rodzaj(ByVal Target As Range, ByVal Arkusz As Worksheet)
    Dim Komorka As Range
    Dim Wiersz As Long

    Application.EnableEvents = False ' Wy��cz zdarzenia, aby unikn�� zap�tlenia

    ' Sprawd� ca�� kolumn� RODZAJ (B)
    For Wiersz = 2 To Arkusz.Cells(Arkusz.Rows.Count, "B").End(xlUp).Row
        If Arkusz.Cells(Wiersz, "B").Value = "Piankowy" Then
            ' Je�li w kolumnie RODZAJ jest "Piankowy", wyczy�� dane w SPRʯYNA i R. SPRʯYNA
            Arkusz.Cells(Wiersz, "D").ClearContents ' Wyczy�� kom�rk� w kolumnie SPRʯYNA (D)
            Arkusz.Cells(Wiersz, "E").ClearContents ' Wyczy�� kom�rk� w kolumnie R. SPRʯYNA (E)
        End If
    Next Wiersz

    Application.EnableEvents = True ' W��cz zdarzenia ponownie
End Sub

