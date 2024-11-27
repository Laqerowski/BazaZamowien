Attribute VB_Name = "KodKokos"
Public Sub Kod_Kokos(ByVal Target As Range, ByVal Arkusz As Worksheet)
    Dim Komorka As Range
    Dim WierszZakresu As Range
    Dim DaneKokos As String
    Dim LewaCzesc As String, PrawaCzesc As String

    ' Zmie� zakres do wszystkich wierszy arkusza
    Set WierszZakresu = Arkusz.Range("B2:B" & Arkusz.Cells(Arkusz.Rows.Count, "B").End(xlUp).Row)

    Application.EnableEvents = False ' Wy��cz zdarzenia, aby unikn�� zap�tlenia

    ' Przejd� przez ka�dy wiersz w arkuszu
    For Each Komorka In WierszZakresu.Columns("B").Cells
        ' Sprawd�, czy RODZAJ (B) i TYP (C) s� wype�nione
        If Komorka.Value <> "" And Arkusz.Cells(Komorka.Row, "C").Value <> "" Then
            ' Je�li kolumna KOKOS (F) jest pusta, wprowad� 0/0
            If Arkusz.Cells(Komorka.Row, "F").Value = "" Then
                Arkusz.Cells(Komorka.Row, "F").Value = "0/0"
            Else
                ' Je�li kolumna KOKOS (F) zawiera dane, sprawd� ich poprawno��
                DaneKokos = Arkusz.Cells(Komorka.Row, "F").Value
                If Not IsValidKokos(DaneKokos) Then
                    Arkusz.Cells(Komorka.Row, "F").Value = "0/0"
                Else
                    ' Dodatkowa walidacja maksymalnej warto�ci 9/9
                    LewaCzesc = Split(DaneKokos, "/")(0)
                    PrawaCzesc = Split(DaneKokos, "/")(1)
                    If Val(LewaCzesc) > 9 Or Val(PrawaCzesc) > 9 Then
                        Arkusz.Cells(Komorka.Row, "F").Value = "0/0"
                    End If
                End If
            End If
        End If
    Next Komorka

    Application.EnableEvents = True ' W��cz zdarzenia ponownie
End Sub

' Funkcja sprawdzaj�ca poprawno�� formatu x/y
Private Function IsValidKokos(ByVal Kokos As String) As Boolean
    Dim Podzielone() As String

    ' Sprawd�, czy dane zawieraj� dok�adnie jeden "/"
    If Len(Kokos) - Len(Replace(Kokos, "/", "")) <> 1 Then
        IsValidKokos = False
        Exit Function
    End If

    ' Podziel dane na cz�ci
    Podzielone = Split(Kokos, "/")
    If UBound(Podzielone) <> 1 Then
        IsValidKokos = False
        Exit Function
    End If

    ' Sprawd�, czy obie cz�ci s� liczbami
    If Not IsNumeric(Podzielone(0)) Or Not IsNumeric(Podzielone(1)) Then
        IsValidKokos = False
        Exit Function
    End If

    ' Je�li wszystko jest w porz�dku, zwr�� True
    IsValidKokos = True
End Function

