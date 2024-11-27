Attribute VB_Name = "KodSpr�yna"
Public Sub Kod_Sprezyna(ByVal Target As Range, ByVal Arkusz As Worksheet)
    Dim Spr�ynaKolumna As Range
    Dim Kom�rka As Range
    
    ' Ustaw zakres kolumny SPRʯYNA (np. kolumna D, zmie� w zale�no�ci od potrzeb)
    Set Spr�ynaKolumna = Intersect(Arkusz.Columns("D"), Target)

    ' Sprawd�, czy zmiana dotyczy kolumny SPRʯYNA
    If Not Spr�ynaKolumna Is Nothing Then
        Application.EnableEvents = False ' Wy��cz zdarzenia, aby unikn�� zap�tlenia

        ' Przejd� przez ka�d� zmienion� kom�rk� w kolumnie SPRʯYNA
        For Each Kom�rka In Spr�ynaKolumna
            If Kom�rka.Value = "Bonel" Or Kom�rka.Value = "MultiPocket" Then
                ' Je�li warto�� to Bonel lub MultiPocket, ustaw Bezstrefowa
                Kom�rka.Offset(0, 1).Value = "Bezstrefowa" ' Przesuni�cie o 1 kolumn� w prawo
            ElseIf Kom�rka.Value = "Kieszeniowa" Or Kom�rka.Value = "Minikiesze�" Then
                ' Je�li warto�� to Kieszeniowa lub Minikiesze�, ustaw domy�ln� warto�� lub list�
                If Kom�rka.Offset(0, 1).Validation.Type <> xlValidateList Then
                    Kom�rka.Offset(0, 1).Validation.Delete ' Usu� istniej�c� walidacj�
                    With Kom�rka.Offset(0, 1).Validation
                        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                            xlBetween, Formula1:="Bezstrefowa,Strefowa"
                        .IgnoreBlank = True
                        .InCellDropdown = True
                    End With
                End If
            Else
                ' Je�li warto�� w SPRʯYNA jest inna, wyczy�� s�siedni� kom�rk�
                Kom�rka.Offset(0, 1).Value = ""
            End If
        Next Kom�rka
        
        Application.EnableEvents = True ' W��cz zdarzenia ponownie
    End If
End Sub


