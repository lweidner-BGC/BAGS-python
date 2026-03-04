Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1
Option Explicit

Private Sub cbAccept_Click()
    If Not IsNumeric(tbQw.Value) Then
        MsgBox "Incorrect or no input!  Reenter discharge before clicking ""Accept"".  " & _
            "Click ""Cancel"" to quit the calculation.", vbOKOnly + vbCritical, "Discharge"
        Exit Sub
    End If
    Me.Hide
    If obCMS.Value Then
        Worksheets("Input").Cells(1, 16).Value = tbQw.Value
    Else
        Worksheets("Input").Cells(1, 16).Value = tbQw.Value * 0.3048 ^ 3
    End If
    Range(Worksheets("Input").Cells(2, 16), Worksheets("Input").Cells(26, 16)).ClearContents
    Unload Me
    UserFormInUse = False
End Sub

Private Sub CommandButton1_Click()
    Me.Hide
    Unload Me
    Canceled = True
    UserFormInUse = False
End Sub

Private Sub obCFS_Click()
    Dim cc As String, Qw As Double
    cc = lbUnit.Caption
    If obCMS.Value Then
        lbUnit.Caption = "cms"
        If lbUnit.Caption <> cc And IsNumeric(tbQw.Value) Then
            Qw = tbQw.Value
            tbQw.Value = Format(Qw * 0.3048 ^ 3, "###0.##")
        End If
    Else
        lbUnit.Caption = "cfs"
        If lbUnit.Caption <> cc And IsNumeric(tbQw.Value) Then
            Qw = tbQw.Value
            tbQw.Value = Format(Qw / 0.3048 ^ 3, "###0.##")
        End If
    End If
End Sub

Private Sub obCMS_Click()
    Dim cc As String, Qw As Double
    cc = lbUnit.Caption
    If obCMS.Value Then
        lbUnit.Caption = "cms"
        If lbUnit.Caption <> cc And IsNumeric(tbQw.Value) Then
            Qw = tbQw.Value
            tbQw.Value = Format(Qw * 0.3048 ^ 3, "###0.##")
        End If
    Else
        lbUnit.Caption = "cfs"
        If lbUnit.Caption <> cc And IsNumeric(tbQw.Value) Then
            Qw = tbQw.Value
            tbQw.Value = Format(Qw / 0.3048 ^ 3, "###0.##")
        End If
    End If
End Sub


