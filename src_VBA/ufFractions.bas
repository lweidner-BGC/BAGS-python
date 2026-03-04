Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1
Option Explicit

Private Sub cbAccept_Click()
    Dim InSt As Worksheet
    Set InSt = Worksheets("Input")
    If IsNumeric(tbGravel.Value) And IsNumeric(tbD65.Value) Then
        InSt.Cells(13, 2).Value = tbD65.Value
        InSt.Cells(12, 1).Value = tbGravel.Value
        Me.Hide
        Unload Me
        UserFormInUse = False
    End If
End Sub

Private Sub CommandButton1_Click()
    Me.Hide
    Unload Me
    Canceled = True
    UserFormInUse = False
End Sub

Private Sub tbGravel_Change()
    If Not IsNumeric(tbGravel.Value) Then Exit Sub
    If tbGravel.Value > 1 Then tbGravel.Value = 1
    If tbGravel.Value < 0 Then tbGravel.Value = 0
    tbSand.Value = 1 - tbGravel.Value
End Sub


