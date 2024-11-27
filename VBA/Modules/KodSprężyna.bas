Attribute VB_Name = "KodSprê¿yna"
Public Sub Kod_Sprezyna(ByVal Target As Range, ByVal Arkusz As Worksheet)
    Dim Sprê¿ynaKolumna As Range
    Dim Komórka As Range
    
    ' Ustaw zakres kolumny SPRÊ¯YNA (np. kolumna D, zmieñ w zale¿noœci od potrzeb)
    Set Sprê¿ynaKolumna = Intersect(Arkusz.Columns("D"), Target)

    ' SprawdŸ, czy zmiana dotyczy kolumny SPRÊ¯YNA
    If Not Sprê¿ynaKolumna Is Nothing Then
        Application.EnableEvents = False ' Wy³¹cz zdarzenia, aby unikn¹æ zapêtlenia

        ' PrzejdŸ przez ka¿d¹ zmienion¹ komórkê w kolumnie SPRÊ¯YNA
        For Each Komórka In Sprê¿ynaKolumna
            If Komórka.Value = "Bonel" Or Komórka.Value = "MultiPocket" Then
                ' Jeœli wartoœæ to Bonel lub MultiPocket, ustaw Bezstrefowa
                Komórka.Offset(0, 1).Value = "Bezstrefowa" ' Przesuniêcie o 1 kolumnê w prawo
            ElseIf Komórka.Value = "Kieszeniowa" Or Komórka.Value = "Minikieszeñ" Then
                ' Jeœli wartoœæ to Kieszeniowa lub Minikieszeñ, ustaw domyœln¹ wartoœæ lub listê
                If Komórka.Offset(0, 1).Validation.Type <> xlValidateList Then
                    Komórka.Offset(0, 1).Validation.Delete ' Usuñ istniej¹c¹ walidacjê
                    With Komórka.Offset(0, 1).Validation
                        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                            xlBetween, Formula1:="Bezstrefowa,Strefowa"
                        .IgnoreBlank = True
                        .InCellDropdown = True
                    End With
                End If
            Else
                ' Jeœli wartoœæ w SPRÊ¯YNA jest inna, wyczyœæ s¹siedni¹ komórkê
                Komórka.Offset(0, 1).Value = ""
            End If
        Next Komórka
        
        Application.EnableEvents = True ' W³¹cz zdarzenia ponownie
    End If
End Sub


