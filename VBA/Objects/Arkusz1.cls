VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Wywo³aj obie funkcje z modu³ów
    Call Kod_Sprezyna(Target, Me)
    Call Kod_Kokos(Target, Me)
    Call Sprawdzam_Rodzaj(Target, Me)
    Call Kod_Lateks(Target, Me)
    Call SprawdzLateks(Me)
End Sub

