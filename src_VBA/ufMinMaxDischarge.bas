Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1
Option Explicit

Private Sub cbAccept_Click()
    Dim i As Long, QwMax As Double, QwMin As Double, dmy As Double
    
    If Not (IsNumeric(tbMinQw.Value) And IsNumeric(tbMaxQw.Value)) Then
        MsgBox "Incorrect input!  Reenter the discharges before click ""Accept"".  " & _
            "Click ""Cancel"" to quit the calculation", vbOKOnly + vbCritical, "Discharge"
        Exit Sub
    End If
    
    QwMin = tbMinQw.Value: QwMax = tbMaxQw.Value
    
    If QwMin > QwMax Then
        MsgBox "Your maximum discharge is smaller then minmum discharge!  Please " & _
            "correct.", vbOKOnly + vbCritical, "Discharge"
        Exit Sub
    End If
    If QwMin <= 0 Then
        MsgBox "You have non-positive discharge!  Discharge must be a positive number.", _
            vbOKOnly + vbCritical, "Discharge"
        Exit Sub
    End If
    Me.Hide
    If obCMS.Value Then
        Worksheets("Input").Cells(1, 16).Value = QwMin
        Worksheets("Input").Cells(26, 16).Value = QwMax
    Else
        Worksheets("Input").Cells(1, 16).Value = QwMin * 0.3048 ^ 3
        Worksheets("Input").Cells(26, 16).Value = QwMax * 0.3048 ^ 3
    End If
    dmy = Log(QwMax / QwMin) / 25
    dmy = Exp(dmy)
    For i = 2 To 25
        Worksheets("Input").Cells(i, 16).Value = Worksheets("Input").Cells(i - 1, 16) * dmy
    Next
    Unload Me
    UserFormInUse = False
End Sub

Private Sub cbCancel_Click()
    Me.Hide
    Unload Me
    UserFormInUse = False
    Canceled = True
End Sub

Private Sub obCFS_Click()
    Dim cc As String, Qw As Double
    cc = lbUnit1.Caption
    If obCMS.Value Then
        lbUnit1.Caption = "cms"
        lbUnit2.Caption = "cms"
        If lbUnit1.Caption <> cc Then
            If IsNumeric(tbMinQw.Value) Then
                Qw = tbMinQw.Value
                tbMinQw.Value = Format(Qw * 0.3048 ^ 3, "###0.##")
            End If
            If IsNumeric(tbMaxQw.Value) Then
                Qw = tbMaxQw.Value
                tbMaxQw.Value = Format(Qw * 0.3048 ^ 3, "###0.##")
            End If
        End If
    Else
        lbUnit1.Caption = "cfs"
        lbUnit2.Caption = "cfs"
        If lbUnit1.Caption <> cc Then
            If IsNumeric(tbMinQw.Value) Then
                Qw = tbMinQw.Value
                tbMinQw.Value = Format(Qw / 0.3048 ^ 3, "###0.##")
            End If
            If IsNumeric(tbMaxQw.Value) Then
                Qw = tbMaxQw.Value
                tbMaxQw.Value = Format(Qw / 0.3048 ^ 3, "###0.##")
            End If
        End If
    End If
End Sub

Private Sub obCMS_Click()
    Dim cc As String, Qw As Double
    cc = lbUnit1.Caption
    If obCMS.Value Then
        lbUnit1.Caption = "cms"
        lbUnit2.Caption = "cms"
        If lbUnit1.Caption <> cc Then
            If IsNumeric(tbMinQw.Value) Then
                Qw = tbMinQw.Value
                tbMinQw.Value = Format(Qw * 0.3048 ^ 3, "###0.##")
            End If
            If IsNumeric(tbMaxQw.Value) Then
                Qw = tbMaxQw.Value
                tbMaxQw.Value = Format(Qw * 0.3048 ^ 3, "###0.##")
            End If
        End If
    Else
        lbUnit1.Caption = "cfs"
        If lbUnit1.Caption <> cc Then
            If IsNumeric(tbMinQw.Value) Then
                Qw = tbMinQw.Value
                tbMinQw.Value = Format(Qw / 0.3048 ^ 3, "###0.##")
            End If
            If IsNumeric(tbMaxQw.Value) Then
                Qw = tbMaxQw.Value
                tbMaxQw.Value = Format(Qw / 0.3048 ^ 3, "###0.##")
            End If
        End If
    End If
End Sub


