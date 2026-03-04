Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

' This module stores common procedures used by other procedures

Sub ShowAndHide(ShowSt As Worksheet, HideSt As Worksheet)
    ShowSt.Visible = True
    ShowSt.Select
    HideSt.Visible = False
End Sub

Sub AuthorViewAll()
    Dim PasWd As String, i As Long
    PasWd = Application.InputBox("Enter developer's password please:")
    If PasWd <> "not4you" Then Exit Sub
    ThisWorkbook.Activate
    ResetMyMenu
    For i = 1 To Sheets.Count
        Sheets(i).Visible = True
    Next
End Sub

Sub ProtectThisSheet(dmy As Integer)
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
Sub UnprotectThisSheet(dmy As Integer)
    ActiveSheet.Unprotect
End Sub

Sub MessageOnWelcome(MyMsg As String)
    If ActiveSheet.Name <> "Welcome" Then Exit Sub
    Application.ScreenUpdating = True
    UnprotectThisSheet 1
    ActiveSheet.Shapes("MyMsg").Visible = True
    ActiveSheet.Shapes("Text Box 13").Visible = True
    ActiveSheet.DrawingObjects("MyMsg").Characters.Text = MyMsg
    Cells(1, 1).Select
    ProtectThisSheet 1
End Sub

Sub EndMessageOnWelcome(dmy As Integer)
    If ActiveSheet.Name <> "Welcome" Then Exit Sub
    UnprotectThisSheet 1
    ActiveSheet.Shapes("MyMsg").Visible = False
    ActiveSheet.Shapes("Text Box 13").Visible = False
    ActiveSheet.Shapes("ProgressBarBackground").Visible = False
    ActiveSheet.Shapes("ProgressBarForeground").Visible = False
    ProtectThisSheet 1
End Sub

Sub ModifyMenu()
    Dim i As Integer
    Dim CheckOut As Label
    Dim BagTool As CommandBarPopup

    If OnErrorOn Then On Error GoTo CheckOut
    
    ResetMyMenu
    MenuBars(xlWorksheet).Menus("&File").MenuItems.Add _
        Caption:="Op&en BAGS Project (.bag)", OnAction:="OpenProject", Before:="&New..."
    MenuBars(xlWorksheet).Menus("&File").MenuItems.Add _
        Caption:="Save BAGS P&roject (.bag)", OnAction:="SaveProject", Before:="&New..."
    MenuBars(xlWorksheet).Menus("&File").MenuItems.Add _
        Caption:="Curre&nt BAGS Project Description", OnAction:="ShowProject", Before:="&New..."
    MenuBars(xlWorksheet).Menus("&File").MenuItems.Add _
        Caption:="BAGS Manual Part &1 (Sediment Transport Primer)", OnAction:="ViewDocument1", Before:="&New..."
    MenuBars(xlWorksheet).Menus("&File").MenuItems.Add _
        Caption:="BAGS Manual Part &2 (Software Instructions)", OnAction:="ViewDocument2", Before:="&New..."
    MenuBars(xlWorksheet).Menus("&File").MenuItems.Add _
        Caption:="-", Before:="&New..."

'    Set BagTool = CommandBars(1).Controls.Add _
'        (Type:=msoControlPopup, Before:=CommandBars(1).Controls("View").Index, Temporary:=True)
'    BagTool.Caption = "&Bag Tools"
    
'    MenuBars(xlWorksheet).Menus("&Bag Tools").MenuItems.Add _
'        Caption:="&Grain size statistics", OnAction:="MyToolGrainSizeStatistics"
'    MenuBars(xlWorksheet).Menus("&Bag Tools").MenuItems.Add _
'        Caption:="&Other tools", OnAction:="OtherTools"
    
    MenuBars(xlWorksheet).Menus("View").Enabled = False
    MenuBars(xlWorksheet).Menus("Insert").Enabled = False
    MenuBars(xlWorksheet).Menus("Format").Enabled = False
    MenuBars(xlWorksheet).Menus("Tools").Enabled = False
    MenuBars(xlWorksheet).Menus("Data").Enabled = False
    MenuBars(xlWorksheet).Menus("Help").Enabled = False
    ThisWorkbook.Application.Caption = "Bedload Transport Equations"
    HideAllToolbars 1
    AddMyToolBar 1
    Exit Sub
    
CheckOut:
    MsgBox "An error occured while ""ModifyMenu"" is executed!  This error may " & _
        "be a result of a slight difference in the MS-Excel software used and " & _
        "will not affect your calculation in anyway." & vbLf & vbLf & _
        "Click OK to continue.", vbOKOnly + vbInformation, "BAGS"
End Sub

Sub ResetMyMenu()
    Dim CheckOut As Label
    
    On Error GoTo CheckOut 'this one must be on all the time
    
    MenuBars(xlWorksheet).Reset
    MenuBars(xlWorksheet).Menus("View").Enabled = True
    MenuBars(xlWorksheet).Menus("Insert").Enabled = True
    MenuBars(xlWorksheet).Menus("Format").Enabled = True
    MenuBars(xlWorksheet).Menus("Tools").Enabled = True
    MenuBars(xlWorksheet).Menus("Data").Enabled = True
    MenuBars(xlWorksheet).Menus("Help").Enabled = True
    Application.Caption = "Microsoft Excel"
    RestoreToolbars 1
    If OnErrorOn Then On Error Resume Next
    Application.CommandBars("MyToolbar").Delete
    Exit Sub
    
CheckOut:
End Sub

Sub HideAllToolbars(dmy As Integer)
    Dim TB As CommandBar
    Dim TBNum As Integer
    Dim TBSheet As Worksheet
    Dim i As Integer
    
    Set TBSheet = Worksheets("Storage")
    Application.ScreenUpdating = False
    
    For i = 1 To 20
        TBSheet.Cells(i, 20).Value = ""
    Next
    
    TBNum = 0
    For Each TB In CommandBars
        If TB.Type = msoBarTypeNormal Then
            If TB.Visible Then
                TBNum = TBNum + 1
                TB.Visible = False
                TBSheet.Cells(TBNum, 20) = TB.Name
            End If
        End If
    Next TB
End Sub

Sub RestoreToolbars(dmy As Integer)
    Dim TBSheet As Worksheet
    Dim MyCell As Object
    Dim i As Integer
    
    Set TBSheet = ThisWorkbook.Worksheets("Storage")

    If TBSheet.Cells(1, 20) = "" Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    If OnErrorOn Then On Error Resume Next
    For Each MyCell In TBSheet.Range("T:T") _
        .SpecialCells(xlCellTypeConstants)
        CommandBars(MyCell.Value).Visible = True
    Next
    For i = 1 To 20
        TBSheet.Cells(i, 20).Value = ""
    Next
    Application.ScreenUpdating = True
End Sub

Sub AddMyToolBar(dmy As Integer)
    Dim MyTBar As CommandBar
    Dim MyBtn
    Dim CheckOut As Label
    
    On Error Resume Next 'this one must be on all the time
    Application.CommandBars("MyToolbar").Delete
    
    If OnErrorOn Then On Error GoTo CheckOut
    Set MyTBar = CommandBars.Add
    With MyTBar
        .Name = "MyToolbar"
        .Position = msoBarTop
        .Visible = True
    End With
    
    Application.CommandBars("MyToolbar").Controls.Add Type:=msoControlComboBox, _
        ID:=1733, Before:=1
    Application.CommandBars("MyToolbar").Controls.Add Type:=msoControlButton, _
        ID:=370, Before:=1
    Application.CommandBars("MyToolbar").Controls.Add Type:=msoControlButton, _
        ID:=22, Before:=1
    Application.CommandBars("MyToolbar").Controls.Add Type:=msoControlButton, _
        ID:=19, Before:=1
    Application.CommandBars("MyToolbar").Controls.Add Type:=msoControlButton, _
        ID:=4, Before:=1
    Application.CommandBars("MyToolbar").Controls.Add Type:=msoControlButton, _
        ID:=2520, Before:=1
    
CheckOut:
End Sub

Private Sub OpenProject()
    Dim MyFile As String
    Dim MyFilter As String
    Dim InSt As Worksheet
    Dim cc As String, MyCC As String, c0 As String
    Dim dmy As Double
    Dim i As Long, j As Long, k As Long
    Dim CheckOut As Label
    
    If OnErrorOn Then On Error GoTo CheckOut
    
    Set InSt = Worksheets("Input")
    
    Application.ScreenUpdating = False
    MyFilter = "BAGS Project File (*.bag),*.bag"
    MyFile = Application.GetOpenFilename(Title:="Open Project", FileFilter:=MyFilter, FilterIndex:=1)
    If UCase(MyFile) = "FALSE" Then
        Exit Sub
    End If
    Worksheets("Storage").Cells(9, 2).Value = MyFile
    
    Close
    Open MyFile For Input As #1
        For i = 1 To 5
            Line Input #1, cc
        Next
        If Left(cc, 1) = """" Then cc = Right(cc, Len(cc) - 1)
        If Right(cc, 1) = """" Then cc = Left(cc, Len(cc) - 1)
        Worksheets("Storage").Cells(10, 2).Value = cc
                
        UserFormInUse = True
        Canceled = False
        ShowProject
        
        Do While UserFormInUse
            DoEvents
        Loop
        
        If Canceled Then Exit Sub
        
        MyCC = ""
        Do While Not EOF(1)
            Line Input #1, cc
            MyCC = MyCC & cc
        Loop
    Close #1
    
    cc = ""
    For i = 1 To Len(MyCC)
        c0 = Mid(MyCC, i, 1)
        If c0 <> """" Then
            cc = cc & c0
        End If
    Next
    
    InSt.Cells.ClearContents
    i = 1
    j = 1
    MyCC = ""
    For k = 1 To Len(cc)
        c0 = Mid(cc, k, 1)
        If c0 = "/" Then
            If MyCC <> "" Then
                InSt.Cells(i, j).Value = MyCC
            End If
            i = i + 1
            MyCC = ""
        ElseIf c0 = "\" Then
            If MyCC <> "" Then
                InSt.Cells(i, j).Value = MyCC
            End If
            i = 1
            j = j + 1
            MyCC = ""
        Else
            MyCC = MyCC & c0
        End If
    Next
    
CheckOut:
    Application.ScreenUpdating = True
End Sub

Private Sub SaveProject()
    SaveMyProject 1
End Sub

Sub SaveMyProject(dmy As Integer)
    Dim MyFile As String
    Dim MyFilter As String
    Dim i As Long, j As Long
    Dim CheckOut As Label
    Dim InSt As Worksheet, MyRange As Range
    Dim Nrow As Long, Ncol As Long, m As Long, n As Long
    Dim cc As String
    
    If OnErrorOn Then On Error GoTo CheckOut
    
    Set InSt = Worksheets("Input"): Set MyRange = InSt.UsedRange
    Nrow = MyRange.Rows.Count
    Ncol = MyRange.Columns.Count
    
    j = vbNo
    Do While j = vbNo
        MyFilter = "BAGS Project File (*.bag),*.bag"
        If IsEmpty(Worksheets("Storage").Cells(9, 2)) Or UCase(Worksheets("Storage").Cells(9, 2).Value) = "N/A" Then
            MyFile = ""
        Else
            MyFile = RightPart(Worksheets("Storage").Cells(9, 2).Value)
        End If
        MyFile = Application.GetSaveAsFilename(InitialFileName:=MyFile, Title:="Save Project", FileFilter:=MyFilter, FilterIndex:=1)
        If UCase(MyFile) = "FALSE" Then
            GoTo CheckOut
        End If
        
        If RightPart(MyFile) = ThisWorkbook.Name Then
            i = MsgBox("Incorrect file name, try again!", vbOKOnly + vbExclamation, "Error in saving project")
            j = vbNo
        Else
            If UCase(Dir(MyFile)) = UCase(RightPart(MyFile)) Then
                j = MsgBox("File " & RightPart(MyFile) & " exists.  Overwrite?", vbYesNo, "Overwrite existing file?")
            Else
                j = vbYes
            End If
        End If
    Loop
    Worksheets("Storage").Cells(9, 2).Value = MyFile
    
    Load ufProject
    ufProject.CommandButton1.Caption = "Save"
    If IsEmpty(Worksheets("Storage").Cells(10, 2).Value) Or Worksheets("Storage").Cells(10, 2).Value = "N/A" Then
        ufProject.TBDescript.Value = "Please enter project description!"
    Else
        ufProject.TBDescript.Value = Worksheets("Storage").Cells(10, 2).Value
    End If
    UserFormInUse = True
    ufProject.Show
    
    Do While UserFormInUse
        DoEvents
    Loop
    
    If Canceled Then Exit Sub
    
    Close
    Open MyFile For Output As #1
        Write #1, "This bag file is generated with BAGS software."
        Write #1, "At no circumstances this file should be edited or modified"
        Write #1, "    outside of BAGS software."
        Write #1, "Version " & VersionNumber
        Write #1, Worksheets("Storage").Cells(10, 2).Value
        
        For j = 1 To Ncol
            cc = ""
            m = 0
            For i = 1 To Nrow
                If IsEmpty(InSt.Cells(i, j)) Then
                    m = m + 1
                Else
                    If m > 0 Then
                        For n = 1 To m
                            cc = cc & "/"
                        Next
                    End If
                    cc = cc & InSt.Cells(i, j).Value & "/"
                    m = 0
                End If
            Next
            cc = cc & "\"
            Write #1, cc
        Next

    Close #1
CheckOut:
    UserFormInUse = False
    Application.ScreenUpdating = True
End Sub

Private Sub ShowProject()
    Dim CheckOut As Label
    
    On Error GoTo CheckOut
    
    Load ufProject
    If Worksheets("Storage").Cells(10, 2).Value <> "" Then
        ufProject.TBDescript.Value = Worksheets("Storage").Cells(10, 2).Value
    Else
        ufProject.TBDescript.Value = "N/A"
    End If
    ufProject.Show
    Exit Sub
    
CheckOut:
    MsgBox "An error occured while ""ShowProject"" is executed!"
End Sub

Private Sub ViewDocument1()
    Dim MyPath As String
    Dim CheckOut As Label
    
    On Error GoTo CheckOut
    MyPath = ThisWorkbook.Path & "\"
    ActiveWorkbook.FollowHyperlink Address:="file:///" & MyPath & "BAGSrpt1.pdf", _
        NewWindow:=True, AddHistory:=False
    Exit Sub
CheckOut:
    MsgBox "Unable to open the document!  Please make sure that you have " & _
        "file BAGSrpt1.pdf (Manual Part 1) copied to the directory where BAGS program is located (i.e., " _
        & MyPath & ").  " & _
        "Please also make sure that you have Adobe Acrobat Reader installed " & _
        "in your computer."
End Sub

Private Sub ViewDocument2()
    Dim MyPath As String
    Dim CheckOut As Label
    
    On Error GoTo CheckOut
    MyPath = ThisWorkbook.Path & "\"
    ActiveWorkbook.FollowHyperlink Address:="file:///" & MyPath & "BAGSrpt2.pdf", _
        NewWindow:=True, AddHistory:=False
    Exit Sub
CheckOut:
    MsgBox "Unable to open the document!  Please make sure that you have " & _
        "file BAGSrpt2.pdf (Manual Part 2) copied to the directory where BAGS program is located (i.e., " _
        & MyPath & ").  " & _
        "Please also make sure that you have Adobe Acrobat Reader installed " & _
        "in your computer."
End Sub

Function RightPart(MyFile As String) As String
    Dim cc As String
    Dim i As Long
    Dim CheckOut As Label
    
    If OnErrorOn Then On Error GoTo CheckOut
        
    RightPart = ""
    i = Len(MyFile)
    cc = Right(MyFile, 1)
    RightPart = cc & RightPart
    Do While cc <> "\" And i > 1
        i = i - 1
        cc = Mid(MyFile, i, 1)
        If cc <> "\" Then
            RightPart = cc & RightPart
        End If
    Loop
    Exit Function
    
CheckOut:
End Function

Sub GetCharacteristicGrainSizeinMM(fRange As Range, dRange As Range, _
    Pct As Double, Dpct As Double)
    
    Dim i As Long
    
    For i = 1 To fRange.Cells.Count - 1
        If (Pct <= fRange.Cells(i + 1).Value And Pct >= fRange.Cells(i).Value) Or _
            (Pct >= fRange.Cells(i + 1).Value And Pct <= fRange.Cells(i).Value) Then
            
            Dpct = Log(dRange.Cells(i)) + _
                (Log(dRange.Cells(i + 1)) - Log(dRange.Cells(i))) / _
                (fRange.Cells(i + 1) - fRange.Cells(i)) * (Pct - fRange.Cells(i))
            Dpct = Exp(Dpct)
            Exit Sub
        End If
    Next
    
End Sub

Sub CalculateSurfaceD65AndGravelSandFractions(dmy As Integer)
    Dim Nsize As Long, i As Long
    Dim D65 As Double, pG As Double, Dmin As Double, Dmax As Double
    Dim InSt As Worksheet
    Dim SizeRange As Range, PctRange As Range
    
    Set InSt = Worksheets("Input")
    
    Nsize = 0
    Do While Not IsEmpty(InSt.Cells(Nsize + 1, 5))
        Nsize = Nsize + 1
    Loop
    
    Set SizeRange = Range(InSt.Cells(1, 5), InSt.Cells(Nsize, 5))
    Set PctRange = Range(InSt.Cells(1, 6), InSt.Cells(Nsize, 6))
    
    Dmin = Application.WorksheetFunction.Min(SizeRange)
    Dmax = Application.WorksheetFunction.Max(SizeRange)
    If Dmin >= 2 Then 'all gravel
        pG = 1
    ElseIf Dmax <= 2 Then 'all sand
        pG = 0
    Else 'interpolation
        For i = 1 To Nsize - 1
            If (SizeRange.Cells(i) <= 2 And SizeRange.Cells(i + 1) >= 2) Or _
                (SizeRange.Cells(i) >= 2 And SizeRange.Cells(i + 1) <= 2) Then
                
                pG = PctRange.Cells(i) + (PctRange.Cells(i + 1) - PctRange.Cells(i)) / _
                    (Log(SizeRange.Cells(i + 1)) - Log(SizeRange.Cells(i))) * _
                    (Log(2) - Log(SizeRange.Cells(i)))
                pG = 1 - pG / 100
                Exit For
            End If
        Next
    End If
    InSt.Cells(12, 1).Value = pG
    
    GetCharacteristicGrainSizeinMM PctRange, SizeRange, 65, D65
    InSt.Cells(13, 2).Value = D65
End Sub

Sub CalculateSubstrateMeanGrainSize(dmy As Integer)
    Dim SizeRange As Range, PctRange As Range, D50 As Double, Nsize As Double
    Dim InSt As Worksheet
    
    Set InSt = Worksheets("Input")
    
    Nsize = 0
    Do While Not IsEmpty(InSt.Cells(Nsize + 1, 7))
        Nsize = Nsize + 1
    Loop
    
    Set SizeRange = Range(InSt.Cells(1, 7), InSt.Cells(Nsize, 7))
    Set PctRange = Range(InSt.Cells(1, 8), InSt.Cells(Nsize, 8))
    
    GetCharacteristicGrainSizeinMM PctRange, SizeRange, 50, D50
    InSt.Cells(13, 1).Value = D50
End Sub

Sub CalculateSubstrateGravelSandFraction(Nsize As Long, pG As Double)
    Dim SizeRange As Range, PctRange As Range
    Dim Dmin As Double, Dmax As Double
    Dim MySt As Worksheet, i As Long
    
    Set MySt = Worksheets("Bedload")
    Set SizeRange = Range(MySt.Cells(21, 4), MySt.Cells(20 + Nsize, 4))
    Set SizeRange = Range(MySt.Cells(21, 5), MySt.Cells(20 + Nsize, 5))

    Dmin = Application.WorksheetFunction.Min(SizeRange)
    Dmax = Application.WorksheetFunction.Max(SizeRange)
    If Dmin >= 2 Then 'all gravel
        pG = 1
    ElseIf Dmax <= 2 Then 'all sand
        pG = 0
    Else 'interpolation
        For i = 1 To Nsize - 1
            If (SizeRange.Cells(i) <= 2 And SizeRange.Cells(i + 1) >= 2) Or _
                (SizeRange.Cells(i) >= 2 And SizeRange.Cells(i + 1) <= 2) Then
                
                pG = PctRange.Cells(i) + (PctRange.Cells(i + 1) - PctRange.Cells(i)) / _
                    (Log(SizeRange.Cells(i + 1)) - Log(SizeRange.Cells(i))) * _
                    (Log(2) - Log(SizeRange.Cells(i)))
                pG = 1 - pG / 100
                Exit For
            End If
        Next
    End If

End Sub

Sub AddCommentsToCell(MyCell As Range, cmt As String)
    MyCell.AddComment
    MyCell.Comment.Text Text:=cmt
End Sub

Function MyMsgBox(Prompt As String, Title As String, ButtonDefault As Long) As Long
    
    Load ufMyMsgBox
    ufMyMsgBox.tbAnswer.Visible = False
    ufMyMsgBox.tbAnswer.Value = ButtonDefault
    ufMyMsgBox.Prompt.Caption = Prompt
    ufMyMsgBox.Caption = Title
    ufMyMsgBox.cbYes.Visible = True
    ufMyMsgBox.cbNo.Visible = True
    ufMyMsgBox.cbCancel.Visible = True
    If ButtonDefault = vbYes Then
        ufMyMsgBox.cbYes.TabIndex = 0
    Else
        ufMyMsgBox.cbNo.TabIndex = 0
    End If
    
    UserFormInUse = True
    ufMyMsgBox.Show
    Do While UserFormInUse
        DoEvents
    Loop
    MyMsgBox = ufMyMsgBox.tbAnswer.Value
    Unload ufMyMsgBox
End Function

Sub GetGrainSizeStatistics(Nsize As Long, MySize As Range, MyFiner As Range, ChD() As Double)
    ' ChD(0) = Dg (mm)
    ' ChD(1) = Geometric standard deviation
    ' ChD(2) = D10 (mm)
    ' ChD(3) = D16 (mm)
    ' ChD(4) = D25 (mm)
    ' ChD(5) = D50 (mm)
    ' ChD(6) = D65 (mm)
    ' ChD(7) = D75 (mm)
    ' ChD(8) = D84 (mm)
    ' ChD(9) = D90 (mm)
    
    Dim j As Long
    Dim Psi(21) As Double, f(21) As Double
    
    For j = 1 To Nsize + 1
        Psi(j) = Log(MySize.Cells(j).Value) / Log(2)
    Next
    For j = 1 To Nsize
        f(j) = Abs(MyFiner.Cells(j + 1).Value - MyFiner.Cells(j).Value) / 100
    Next
    
    GetGeometricMeanGrainSizeAndArithmeticStandardDeviation Nsize, Psi, f, ChD(0), ChD(1)
    ChD(0) = ChD(0) * 1000
    ChD(1) = 2 ^ ChD(1)
    
    GetCharacteristicGrainSizeinMM MyFiner, MySize, 10, ChD(2)
    GetCharacteristicGrainSizeinMM MyFiner, MySize, 16, ChD(3)
    GetCharacteristicGrainSizeinMM MyFiner, MySize, 25, ChD(4)
    GetCharacteristicGrainSizeinMM MyFiner, MySize, 50, ChD(5)
    GetCharacteristicGrainSizeinMM MyFiner, MySize, 65, ChD(6)
    GetCharacteristicGrainSizeinMM MyFiner, MySize, 75, ChD(7)
    GetCharacteristicGrainSizeinMM MyFiner, MySize, 84, ChD(8)
    GetCharacteristicGrainSizeinMM MyFiner, MySize, 90, ChD(9)
    
End Sub



