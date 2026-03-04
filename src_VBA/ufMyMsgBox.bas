Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1
Option Explicit

Private Sub cbCancel_Click()
    Me.Hide
    tbAnswer = vbCancel
    UserFormInUse = False
End Sub

Private Sub cbNo_Click()
    Me.Hide
    tbAnswer = vbNo
    UserFormInUse = False
End Sub

Private Sub cbYes_Click()
    Me.Hide
    tbAnswer = vbYes
    UserFormInUse = False
End Sub


