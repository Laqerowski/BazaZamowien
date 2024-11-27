Attribute VB_Name = "KodLateks"
Public Sub Kod_Lateks(ByVal Target As Range, ByVal Arkusz As Worksheet)
    Dim Komorka As Range
    Dim WierszZakresu As Range
    Dim DaneLateks As String
    Dim LewaCzesc As String, PrawaCzesc As String

    ' Ustaw zakres dla kolumny B (RODZAJ)
    Set WierszZakresu = Arkusz.Range("B2:B" & Arkusz.Cells(Arkusz.Rows.Count, "B").End(xlUp).Row)

    Application.EnableEvents = False ' Wy��cz zdarzenia, aby unikn�� zap�tlenia

    ' Przejd� przez ka�dy wiersz w arkuszu
    For Each Komorka In WierszZakresu.Cells
        ' Sprawd�, czy RODZAJ (B) i TYP (C) s� wype�nione
        If Komorka.Value <> "" And Arkusz.Cells(Komorka.Row, "C").Value <> "" Then
            ' Je�li kolumna LATEKS (G) jest pusta, wprowad� 0/0
            If Arkusz.Cells(Komorka.Row, "G").Value = "" Then
                Arkusz.Cells(Komorka.Row, "G").Value = "0/0"
            Else
                ' Je�li kolumna LATEKS (G) zawiera dane, sprawd� ich poprawno��
                DaneLateks = Arkusz.Cells(Komorka.Row, "G").Value
                If Not IsValidLateks(DaneLateks) Then
                    Arkusz.Cells(Komorka.Row, "G").Value = "0/0"
                Else
                    ' Dodatkowa walidacja maksymalnej warto�ci 9/9
                    LewaCzesc = Split(DaneLateks, "/")(0)
                    PrawaCzesc = Split(DaneLateks, "/")(1)
                    If Val(LewaCzesc) > 9 Or Val(PrawaCzesc) > 9 Then
                        Arkusz.Cells(Komorka.Row, "G").Value = "0/0"
                    End If
                End If
            End If
        End If
    Next Komorka

    Application.EnableEvents = True ' W��cz zdarzenia ponownie
End Sub

' Funkcja sprawdzaj�ca poprawno�� formatu x/y
Private Function IsValidLateks(ByVal Lateks As String) As Boolean
    Dim Podzielone() As String

    ' Sprawd�, czy dane zawieraj� dok�adnie jeden "/"
    If Len(Lateks) - Len(Replace(Lateks, "/", "")) <> 1 Then
        IsValidLateks = False
        Exit Function
    End If

    ' Podziel dane na cz�ci
    Podzielone = Split(Lateks, "/")
    If UBound(Podzielone) <> 1 Then
        IsValidLateks = False
        Exit Function
    End If

    ' Sprawd�, czy obie cz�ci s� liczbami
    If Not IsNumeric(Podzielone(0)) Or Not IsNumeric(Podzielone(1)) Then
        IsValidLateks = False
        Exit Function
    End If

    ' Je�li wszystko jest w porz�dku, zwr�� True
    IsValidLateks = True
End Function

