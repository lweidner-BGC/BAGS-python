Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1
Option Explicit

Private Sub cbAccept_Click()
    Dim InSt As Worksheet
    Set InSt = Worksheets("Input")
    If obWS Then InSt.Cells(5, 1).Value = "W.S."
    If obBed Then InSt.Cells(5, 1).Value = "Bed"
    If obModel Then InSt.Cells(5, 1).Value = "Model"
    InSt.Cells(5, 2).Value = tbSlope.Value
    UserFormInUse = False
    Me.Hide
    Unload Me
End Sub

Private Sub cbCancel_Click()
    UserFormInUse = False
    Canceled = True
    Me.Hide
    Unload Me
End Sub


