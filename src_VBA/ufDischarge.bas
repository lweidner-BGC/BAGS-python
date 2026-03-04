Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1
Option Explicit

Private Sub cbAccept_Click()
    If obA Then
        Worksheets("Input").Cells(6, 1).Value = "(A)"
    ElseIf obB Then
        Worksheets("Input").Cells(6, 1).Value = "(B)"
    ElseIf obC Then
        If obDuration Then
            Worksheets("Input").Cells(6, 1).Value = "(C1)"
        ElseIf obRecord Then
            Worksheets("Input").Cells(6, 1).Value = "(C2)"
        Else
            MsgBox "You must select whether to use discharge record or flow " & _
                "duration curve as input.", vbOKOnly + vbCritical, "Missing information"
            Exit Sub
        End If
    End If
    UserFormInUse = False
    Me.Hide
    Unload Me
End Sub

Private Sub CommandButton1_Click()
    Me.Hide
    Unload Me
    UserFormInUse = False
    Canceled = True
End Sub

Private Sub obA_Click()
    If obC Then
        Frame1.Visible = True
        If Worksheets("Input").Cells(6, 1).Value = "(C1)" Then obDuration.Value = True
        If Worksheets("Input").Cells(6, 1).Value = "(C2)" Then obRecord.Value = True
    Else
        Frame1.Visible = False
    End If
End Sub

Private Sub obB_Click()
    If obC Then
        Frame1.Visible = True
        If Worksheets("Input").Cells(6, 1).Value = "(C1)" Then obDuration.Value = True
        If Worksheets("Input").Cells(6, 1).Value = "(C2)" Then obRecord.Value = True
    Else
        Frame1.Visible = False
    End If
End Sub

Private Sub obC_Click()
    If obC Then
        Frame1.Visible = True
        If Worksheets("Input").Cells(6, 1).Value = "(C1)" Then obDuration.Value = True
        If Worksheets("Input").Cells(6, 1).Value = "(C2)" Then obRecord.Value = True
    Else
        Frame1.Visible = False
    End If
End Sub


