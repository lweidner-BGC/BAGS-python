Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column = 7 Then
        PlotFloodplains 1
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    UnprotectThisSheet 1
    Range(Cells(23, 6), Cells(24, 7)).Interior.ColorIndex = 2
    Target.Interior.ColorIndex = 15
    ProtectThisSheet 1
End Sub


