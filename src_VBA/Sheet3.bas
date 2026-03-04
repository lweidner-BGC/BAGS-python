Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Option Explicit

Private Sub Worksheet_Activate()
    ActiveSheet.Shapes("obCMS").Visible = False
    ActiveSheet.Shapes("obCFS").Visible = False
End Sub



