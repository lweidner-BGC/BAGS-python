Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1
Option Explicit

Private Sub cbAccept_Click()
    Option1 = OptionButton1
    Option2 = OptionButton2
    If Option1 Or Option2 Then
        UserFormInUse = False
        Me.Hide
        Unload Me
    Else
        MsgBox "You must make an option in order to continue, or click " & _
            """Cancel"" to quit the program."
    End If
End Sub

Private Sub cbCancel_Click()
    Me.Hide
    Option1 = False
    Option2 = False
    UserFormInUse = False
    Unload Me
End Sub

Private Sub OptionButton1_Click()

End Sub


