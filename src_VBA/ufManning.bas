Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1


Option Explicit

Private Sub CommandButton1_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub UserForm_Activate()
    ufManning.ScrollTop = 0
End Sub



