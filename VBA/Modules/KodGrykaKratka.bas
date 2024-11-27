Attribute VB_Name = "KodGrykaKratka"
Public Sub Kod_GrykaKratka(ByVal Target As Range, ByVal Arkusz As Worksheet)
    Dim RodzajKolumna As Range
    Dim Komórka As Range
    
    ' Ustaw zakres kolumny RODZAJ (np. kolumna B, zmieñ w zale¿noœci od potrzeb)
    Set RodzajKolumna = Intersect(Arkusz.Columns("B"), Target)

    ' SprawdŸ, czy zmiana dotyczy kolumny RODZAJ
    If Not RodzajKolumna Is Nothing Then
        Application.EnableEvents = False ' Wy³¹cz zdarzenia, aby unikn¹æ zapêtlenia

        ' PrzejdŸ przez ka¿d¹ zmienion¹ komórkê w kolumnie RODZAJ
        For Each Komórka In RodzajKolumna
            If Komórka.Value = "Profilowana" Or Komórka.Value = "Gryka sypana" Then
                ' Jeœli wartoœæ to Profilowana lub Gryka sypana, ustaw NIE
                Komórka.Offset(0, 1).Value = "NIE" ' Przesuniêcie o 1 kolumnê w prawo
            ElseIf Komórka.Value = "Piankowa" Or Komórka.Value = "Kulka silikonowa" Or Komórka.Value = "Lateksowa" Then
                ' Jeœli wartoœæ to Piankowa, Kulka silikonowa lub Lateksowa ustaw domyœln¹ wartoœæ lub listê
                If Komórka.Offset(0, 1).Validation.Type <> xlValidateList Then
                    Komórka.Offset(0, 1).Validation.Delete ' Usuñ istniej¹c¹ walidacjê
                    With Komórka.Offset(0, 1).Validation
                        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                            xlBetween, Formula1:="TAK,NIE"
                        .IgnoreBlank = True
                        .InCellDropdown = True
                    End With
                End If
            Else
                ' Jeœli wartoœæ w RODZAJ jest inna, wyczyœæ s¹siedni¹ komórkê
                Komórka.Offset(0, 1).Value = ""
            End If
        Next Komórka
        
        Application.EnableEvents = True ' W³¹cz zdarzenia ponownie
    End If
End Sub


