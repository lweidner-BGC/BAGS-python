Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Option Explicit

Private Sub Worksheet_Activate()
    ActiveSheet.ScrollArea = "A1"
    ActiveSheet.Shapes("MyMsg").Visible = False
    ProtectThisSheet 1
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = Cells(1, 1).Address Then
        If UCase(Cells(1, 1).Value) = "AUTHOR" Then
            MsgBox "Dr. Yantao Cui is a hydraulic engineer, " & _
                "specializing in open channel hydraulics and sediment transport.  " & _
                "He received his Ph.D. in civil engineering in 1996 from the " & _
                "University of Minnesota." & vbLf & vbLf & "Contact Yantao by Email at " & _
                "ytc@astound.net if you would like him to be involved in " & _
                "a project.", _
                vbOKOnly + vbInformation, "About the author"
        End If
        Cells(1, 1).ClearContents
    End If
End Sub


