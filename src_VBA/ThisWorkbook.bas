Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Option Explicit

Private Sub Workbook_Open()
    Dim i As Long
    
    On Error Resume Next
    
'    If Year(Date) > 2004 Then
'        MsgBox "This version of BAGS software is becoming obsolete.  " & _
'            "Please check into STREAM Systems Technology Center's home page at " & _
'            "http://www.stream.fs.fed.us/ " & _
'            "to download the latest version." & vbLf & vbLf & _
'            "Thank you for evaluating BAGS software and good luck with the new version.", _
'            vbOKOnly + vbInformation, "BAGS"
'        ThisWorkbook.Close
'        Exit Sub
'    End If

'    i = MsgBox("Six bedload equations available in literature are implemented in this software.  " & _
'        "It is possible that there are mistakes in the implementation of the equations.  " & _
'        "Use the software with your own judgement and at your own risk.  Neither the Forest Service nor the authors are " & _
'        "responsible for the damages resulted from the application of this software." & vbLf & vbLf & _
'        "Agree? (Click Yes to continue or No to close the program.)", _
'        vbYesNo + vbQuestion, "Application Agreement")
    
'    If i = vbNo Then
'        ThisWorkbook.Close
'        Exit Sub
'    End If
    
    'ufSplashScreen.Show
    Application.ScreenUpdating = False
    Worksheets("Storage").Cells.ClearContents
    Application.OnWindow = "ResetMyMenu"
    Application.Windows(ThisWorkbook.Name).OnWindow = "ModifyMenu"
    Application.ScreenUpdating = True
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    Application.OnWindow = "ResetMyMenu"
    Application.Windows(ThisWorkbook.Name).OnWindow = "ModifyMenu"
End Sub

Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
    On Error Resume Next
    ResetMyMenu
    Application.OnWindow = ""
End Sub


