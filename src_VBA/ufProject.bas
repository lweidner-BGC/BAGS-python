Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1


Option Explicit

Private Sub CommandButton1_Click()
    Dim CheckOut As Label
    
    On Error GoTo CheckOut
    
    Worksheets("Storage").Cells(10, 2).Value = Left(TBDescript.Value, 1010)
    ufProject.Hide
    
    Unload ufProject
    UserFormInUse = False
    Exit Sub
    
CheckOut:
    UserFormInUse = False
    MsgBox "An error occured while saving project description!"
End Sub

Private Sub CommandButton2_Click()
    Dim CheckOut As Label
    
    On Error GoTo CheckOut
    
    ufProject.Hide
    Unload ufProject
    UserFormInUse = False
    Canceled = True
    Exit Sub
    
CheckOut:
    UserFormInUse = False
    MsgBox "An error occured while exiting project description!"
End Sub



