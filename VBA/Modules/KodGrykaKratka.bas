Attribute VB_Name = "KodGrykaKratka"
Public Sub Kod_GrykaKratka(ByVal Target As Range, ByVal Arkusz As Worksheet)
    Dim RodzajKolumna As Range
    Dim Kom�rka As Range
    
    ' Ustaw zakres kolumny RODZAJ (np. kolumna B, zmie� w zale�no�ci od potrzeb)
    Set RodzajKolumna = Intersect(Arkusz.Columns("B"), Target)

    ' Sprawd�, czy zmiana dotyczy kolumny RODZAJ
    If Not RodzajKolumna Is Nothing Then
        Application.EnableEvents = False ' Wy��cz zdarzenia, aby unikn�� zap�tlenia

        ' Przejd� przez ka�d� zmienion� kom�rk� w kolumnie RODZAJ
        For Each Kom�rka In RodzajKolumna
            If Kom�rka.Value = "Profilowana" Or Kom�rka.Value = "Gryka sypana" Then
                ' Je�li warto�� to Profilowana lub Gryka sypana, ustaw NIE
                Kom�rka.Offset(0, 1).Value = "NIE" ' Przesuni�cie o 1 kolumn� w prawo
            ElseIf Kom�rka.Value = "Piankowa" Or Kom�rka.Value = "Kulka silikonowa" Or Kom�rka.Value = "Lateksowa" Then
                ' Je�li warto�� to Piankowa, Kulka silikonowa lub Lateksowa ustaw domy�ln� warto�� lub list�
                If Kom�rka.Offset(0, 1).Validation.Type <> xlValidateList Then
                    Kom�rka.Offset(0, 1).Validation.Delete ' Usu� istniej�c� walidacj�
                    With Kom�rka.Offset(0, 1).Validation
                        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                            xlBetween, Formula1:="TAK,NIE"
                        .IgnoreBlank = True
                        .InCellDropdown = True
                    End With
                End If
            Else
                ' Je�li warto�� w RODZAJ jest inna, wyczy�� s�siedni� kom�rk�
                Kom�rka.Offset(0, 1).Value = ""
            End If
        Next Kom�rka
        
        Application.EnableEvents = True ' W��cz zdarzenia ponownie
    End If
End Sub


