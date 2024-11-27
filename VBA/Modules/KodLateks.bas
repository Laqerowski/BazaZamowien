Attribute VB_Name = "KodLateks"
Public Sub Kod_Lateks(ByVal Target As Range, ByVal Arkusz As Worksheet)
    Dim Komorka As Range
    Dim WierszZakresu As Range
    Dim DaneLateks As String
    Dim LewaCzesc As String, PrawaCzesc As String

    ' Ustaw zakres dla kolumny B (RODZAJ)
    Set WierszZakresu = Arkusz.Range("B2:B" & Arkusz.Cells(Arkusz.Rows.Count, "B").End(xlUp).Row)

    Application.EnableEvents = False ' Wy³¹cz zdarzenia, aby unikn¹æ zapêtlenia

    ' PrzejdŸ przez ka¿dy wiersz w arkuszu
    For Each Komorka In WierszZakresu.Cells
        ' SprawdŸ, czy RODZAJ (B) i TYP (C) s¹ wype³nione
        If Komorka.Value <> "" And Arkusz.Cells(Komorka.Row, "C").Value <> "" Then
            ' Jeœli kolumna LATEKS (G) jest pusta, wprowadŸ 0/0
            If Arkusz.Cells(Komorka.Row, "G").Value = "" Then
                Arkusz.Cells(Komorka.Row, "G").Value = "0/0"
            Else
                ' Jeœli kolumna LATEKS (G) zawiera dane, sprawdŸ ich poprawnoœæ
                DaneLateks = Arkusz.Cells(Komorka.Row, "G").Value
                If Not IsValidLateks(DaneLateks) Then
                    Arkusz.Cells(Komorka.Row, "G").Value = "0/0"
                Else
                    ' Dodatkowa walidacja maksymalnej wartoœci 9/9
                    LewaCzesc = Split(DaneLateks, "/")(0)
                    PrawaCzesc = Split(DaneLateks, "/")(1)
                    If Val(LewaCzesc) > 9 Or Val(PrawaCzesc) > 9 Then
                        Arkusz.Cells(Komorka.Row, "G").Value = "0/0"
                    End If
                End If
            End If
        End If
    Next Komorka

    Application.EnableEvents = True ' W³¹cz zdarzenia ponownie
End Sub

' Funkcja sprawdzaj¹ca poprawnoœæ formatu x/y
Private Function IsValidLateks(ByVal Lateks As String) As Boolean
    Dim Podzielone() As String

    ' SprawdŸ, czy dane zawieraj¹ dok³adnie jeden "/"
    If Len(Lateks) - Len(Replace(Lateks, "/", "")) <> 1 Then
        IsValidLateks = False
        Exit Function
    End If

    ' Podziel dane na czêœci
    Podzielone = Split(Lateks, "/")
    If UBound(Podzielone) <> 1 Then
        IsValidLateks = False
        Exit Function
    End If

    ' SprawdŸ, czy obie czêœci s¹ liczbami
    If Not IsNumeric(Podzielone(0)) Or Not IsNumeric(Podzielone(1)) Then
        IsValidLateks = False
        Exit Function
    End If

    ' Jeœli wszystko jest w porz¹dku, zwróæ True
    IsValidLateks = True
End Function

