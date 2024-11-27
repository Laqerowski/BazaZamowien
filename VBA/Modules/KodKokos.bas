Attribute VB_Name = "KodKokos"
Public Sub Kod_Kokos(ByVal Target As Range, ByVal Arkusz As Worksheet)
    Dim Komorka As Range
    Dim WierszZakresu As Range
    Dim DaneKokos As String
    Dim LewaCzesc As String, PrawaCzesc As String

    ' Zmieñ zakres do wszystkich wierszy arkusza
    Set WierszZakresu = Arkusz.Range("B2:B" & Arkusz.Cells(Arkusz.Rows.Count, "B").End(xlUp).Row)

    Application.EnableEvents = False ' Wy³¹cz zdarzenia, aby unikn¹æ zapêtlenia

    ' PrzejdŸ przez ka¿dy wiersz w arkuszu
    For Each Komorka In WierszZakresu.Columns("B").Cells
        ' SprawdŸ, czy RODZAJ (B) i TYP (C) s¹ wype³nione
        If Komorka.Value <> "" And Arkusz.Cells(Komorka.Row, "C").Value <> "" Then
            ' Jeœli kolumna KOKOS (F) jest pusta, wprowadŸ 0/0
            If Arkusz.Cells(Komorka.Row, "F").Value = "" Then
                Arkusz.Cells(Komorka.Row, "F").Value = "0/0"
            Else
                ' Jeœli kolumna KOKOS (F) zawiera dane, sprawdŸ ich poprawnoœæ
                DaneKokos = Arkusz.Cells(Komorka.Row, "F").Value
                If Not IsValidKokos(DaneKokos) Then
                    Arkusz.Cells(Komorka.Row, "F").Value = "0/0"
                Else
                    ' Dodatkowa walidacja maksymalnej wartoœci 9/9
                    LewaCzesc = Split(DaneKokos, "/")(0)
                    PrawaCzesc = Split(DaneKokos, "/")(1)
                    If Val(LewaCzesc) > 9 Or Val(PrawaCzesc) > 9 Then
                        Arkusz.Cells(Komorka.Row, "F").Value = "0/0"
                    End If
                End If
            End If
        End If
    Next Komorka

    Application.EnableEvents = True ' W³¹cz zdarzenia ponownie
End Sub

' Funkcja sprawdzaj¹ca poprawnoœæ formatu x/y
Private Function IsValidKokos(ByVal Kokos As String) As Boolean
    Dim Podzielone() As String

    ' SprawdŸ, czy dane zawieraj¹ dok³adnie jeden "/"
    If Len(Kokos) - Len(Replace(Kokos, "/", "")) <> 1 Then
        IsValidKokos = False
        Exit Function
    End If

    ' Podziel dane na czêœci
    Podzielone = Split(Kokos, "/")
    If UBound(Podzielone) <> 1 Then
        IsValidKokos = False
        Exit Function
    End If

    ' SprawdŸ, czy obie czêœci s¹ liczbami
    If Not IsNumeric(Podzielone(0)) Or Not IsNumeric(Podzielone(1)) Then
        IsValidKokos = False
        Exit Function
    End If

    ' Jeœli wszystko jest w porz¹dku, zwróæ True
    IsValidKokos = True
End Function

