Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = Cells(16, 5).Address Then
        If Cells(16, 5).Value > 1 Then
            Cells(16, 5) = 1
        ElseIf Cells(16, 5).Value < 0 Then
            Cells(16, 5).Value = 0
        End If
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Address = Cells(18, 5).Address Then
        Cells(20, 5).Select
    ElseIf Target.Address = Cells(19, 5).Address Then
        Cells(17, 5).Select
    End If
End Sub


