VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Wywo³aj funkcjê Kod_Rodzaj
    Call Kod_GrykaKratka(Target, Me)
    Call SprawdzamGrykaKratka(Target, Me)
End Sub

