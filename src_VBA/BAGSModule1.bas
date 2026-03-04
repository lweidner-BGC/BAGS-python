Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

' This module stores the procedures for data input

Public VersionNumber As String
Public LastUpdated As String
Public Parker90 As Boolean, Parker82 As Boolean, PK82 As Boolean
Public Wilcock As Boolean, Wilcock03 As Boolean, Bakke As Boolean
Public Option1 As Boolean, Option2 As Boolean, UserFormInUse As Boolean
Public Canceled As Boolean, OnErrorOn As Boolean
Public EnableOnError As Boolean
Public PKorPKM As String 'to use Parker-Klingeman (1982) or Parker-Klingeman-McLean (1982) for Bakke (1999)

Sub Auto_Open()
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim i As Long, MyTime As Double
    EnableOnError = True
    Worksheets("Welcome").Visible = True
    Worksheets("Welcome").Select
    Application.ScreenUpdating = True
    Worksheets("Welcome").ScrollArea = "A1"
    Worksheets("Agreement").ScrollArea = "A1"
    For i = 1 To Sheets.Count
        If Sheets(i).Name <> "Welcome" Then
            Sheets(i).Visible = False
        End If
    Next
    OnErrorOn = False
    EndMessageOnWelcome 1
    VersionNumber = "2008.11"
    LastUpdated = "November 2008"
End Sub

Sub RefreshPage()
    On Error Resume Next
    Dim i As Long, MyTime As Double
    EnableOnError = True
    Worksheets("Welcome").Visible = True
    Worksheets("Welcome").Select
    Worksheets("Welcome").ScrollArea = "A1"
    Worksheets("Agreement").ScrollArea = "A1"
    For i = 1 To Sheets.Count
        If Sheets(i).Name <> "Agreement" Then
            Sheets(i).Visible = False
        End If
    Next
    OnErrorOn = False
    EndMessageOnWelcome 1
    VersionNumber = "2008.2"
    LastUpdated = "February 2008"
End Sub

Sub Auto_close()
    If EnableOnError Then On Error Resume Next
    ResetMyMenu
    Application.OnWindow = ""
    Worksheets("Welcome").Visible = True
    Worksheets("Welcome").Select
    Application.ThisWorkbook.Save
End Sub

Sub CloseSplashScreen()
    Unload ufSplashScreen
End Sub

Function WindowOS() As Boolean
    If EnableOnError Then On Error Resume Next
    If Application.OperatingSystem Like "*Win*" Then
        WindowOS = True
    Else
        WindowOS = False
    End If
End Function

Sub StartTheApplication()
    Worksheets("Agreement").Visible = True
    Worksheets("Agreement").Select
    Worksheets("Welcome").Visible = False
End Sub

Sub RunSoftware(dd As Integer)
    Dim StInput As Worksheet
    Dim i As Long, j As Long, k As Long
    Dim dmy(10) As Double
    
    If EnableOnError Then On Error Resume Next
    
    Set StInput = Worksheets("Input")
    
    ' select equations to use
    UserFormInUse = True
    Load ufEquations
    ufEquations.cbParker90 = StInput.Cells(8, 2).Value
    ufEquations.cbParker82 = StInput.Cells(9, 2).Value
    ufEquations.cbPK82 = StInput.Cells(8, 1).Value
    ufEquations.cbWilcock = StInput.Cells(10, 2).Value
    ufEquations.cbWilcock03 = StInput.Cells(10, 1).Value
    ufEquations.cbBakke = StInput.Cells(11, 2).Value
    ufEquations.Show
    Do While UserFormInUse
        DoEvents
    Loop
    If Not (Parker90 Or Parker82 Or PK82 Or Wilcock Or Wilcock03 Or Bakke) Then Exit Sub
    
    'option on channel geometry (cross section or bankfull width)
    UserFormInUse = True
    Load ufQuestions
    ufQuestions.Caption = "Channel Geometry"
    ufQuestions.Label1.Caption = _
        "Will you use a typical cross section or reach average bankfull " & _
        "width to represent the channel geometry?"
    If StInput.Cells(1, 1).Value = "XS" Then
        ufQuestions.OptionButton1.Value = True
    Else
        ufQuestions.OptionButton2.Value = True
    End If
    ufQuestions.OptionButton1.Caption = " A typical cross section"
    ufQuestions.OptionButton2.Caption = " Reach average bankfull width"
    ufQuestions.Show
    Do While UserFormInUse
        DoEvents
    Loop
    If (Not Option1) And (Not Option2) Then Exit Sub
    
    If Option1 Then 'go get the cross section
        StInput.Cells(1, 1).Value = "XS"
        StInput.Cells(1, 2).ClearContents
        UserFormInUse = True
        GoGetTheCrossSection
    End If
    
    If Option2 Then 'go get the bankfull width
        StInput.Cells(1, 1).Value = "Width"
        dmy(0) = Application.InputBox("Enter reach average bankfull width, in meters:", _
            "Bankfull Width", StInput.Cells(1, 2).Value)
        If dmy(0) = 0 Then Exit Sub 'canceled
        StInput.Cells(1, 2).Value = dmy(0)
        StInput.Columns("C:D").ClearContents
        GoGetGrainSizeInformation 1
    End If
    
End Sub

Sub AboutTheSoftware()
    If EnableOnError Then On Error Resume Next
    MsgBox "This software implements the following bedload transport equations:" & vbLf & _
        vbLf & _
        "    Parker (1990) surface-based;" & vbLf & _
        "    Parker, Klingeman, and McLean (1982) substrate D50-based;" & vbLf & _
        "    Parker and Klingeman (1982) substrate-based;" & vbLf & _
        "    Wilcock (2001) two-fraction;" & vbLf & _
        "    Wilcock and Crowe (2003) surface-based;" & vbLf & _
        "    Bakke et al. (1999)." & vbLf & vbLf & _
        "Major Technical Contributors: John Pitlick, Peter Wilcock, John Potyondy, " & _
        "Paul Bakke, and Yantao Cui." & vbLf & vbLf & _
        "Dr. Yantao Cui is responsible for the design and coding of this software and can be" & _
        " contacted at ytc@astound.net if you wish to report potential errors within the" & _
        " software or pass along the desired improvements you would like to see in the future." & _
        vbLf & vbLf & "This version (" & VersionNumber & ") was last updated in " & LastUpdated & ".", _
        vbOKOnly + vbInformation, "About BAGS - Version " & VersionNumber
End Sub

Private Sub GoGetTheCrossSection()
    Dim i As Long
    If EnableOnError Then On Error Resume Next
    ShowAndHide Worksheets("MyInput"), Worksheets("Welcome")
    Cells(2, 1).Select
    ActiveSheet.ScrollArea = "A2:B" & ActiveSheet.Rows.Count
    Cells(1, 1).Value = "Distance (m)"
    Cells(1, 2).Value = "Elevation (m)"
    Cells(3, 4).Value = "Enter cross section data in columns A and B, " & _
        "with column A being lateral distance and column B being elevation.  " & _
        "The elevation can be relative to an arbitrary datum." & vbLf & vbLf & _
        "Note that you can copy data directly to the cells from another workbook."
    ClearFormClickedOnMyInput
    If Not IsEmpty(Worksheets("Input").Cells(1, 3)) Then
        i = 0
        Do While Not IsEmpty(Worksheets("Input").Cells(i + 1, 3))
            i = i + 1
            Cells(i + 1, 1).Value = Worksheets("Input").Cells(i, 3).Value
            Cells(i + 1, 2).Value = Worksheets("Input").Cells(i, 4).Value
        Loop
    End If
End Sub

Sub AcceptClickedOnMyInput()
    Dim i As Long, j As Long, k As Long, QwUnit As Double, dmy As Double
    
    If EnableOnError Then On Error Resume Next
    If IsEmpty(Cells(2, 1)) Then
        MsgBox "Input data must start at the 2nd row!", vbOKOnly + vbCritical, "Incorrect input"
        Exit Sub
    End If
        
    If Cells(1, 1).Value = "Distance (m)" Then 'input for cross section
        i = 1
        Do While Not IsEmpty(Cells(i + 1, 1))
            i = i + 1
            Worksheets("Input").Cells(i - 1, 3).Value = Cells(i, 1).Value
            Worksheets("Input").Cells(i - 1, 4).Value = Cells(i, 2).Value
        Loop
        Range(Worksheets("Input").Cells(i, 3), _
            Worksheets("Input").Cells(Worksheets("Input").Rows.Count, 4)).ClearContents
        ShowAndHide Worksheets("Welcome"), Worksheets("MyInput")
            
        ShowCrossSectionAndGetFloodplainInformation
        Exit Sub
    End If
    
    If Cells(1, 1).Value = "Size (mm)" And Left(Cells(3, 4).Value, 3) = "Sur" Then
    'surface grain size distribution
        i = 1
        Do While Not IsEmpty(Cells(i + 1, 1))
            i = i + 1
            Worksheets("Input").Cells(i - 1, 5).Value = Cells(i, 1).Value
            Worksheets("Input").Cells(i - 1, 6).Value = Cells(i, 2).Value
        Loop
        Range(Worksheets("Input").Cells(i, 5), _
            Worksheets("Input").Cells(Worksheets("Input").Rows.Count, 6)).ClearContents
        ShowAndHide Worksheets("Welcome"), Worksheets("MyInput")
        
        If Parker90 Then ' check if the finest grain size is coarser than 2 mm
            dmy = Application.WorksheetFunction.Min( _
                Range(Worksheets("Input").Cells(1, 5), Worksheets("Input").Cells(i - 1, 5)))
            If dmy >= 2 Then
                Worksheets("Storage").Cells(1, 1).Value = -1
            Else ' check out whether need adjustment in grain size distribution
                i = MyMsgBox("You have chosen to apply the surface-based bedload equation of " & _
                    "Parker (1990), which is suggested to be applicable to particles " & _
                    "coarser then 2 mm.  There are particles finer than 2 mm in your " & _
                    "surface grain size.  In applying Parker's equation, you can either " & _
                    "force the program to run for given grain size distribution, or to allow the program " & _
                    "to ignore the finer particles." & vbLf & vbLf & "Ignore " & _
                    "the finer particles?", "Surface grain size", Worksheets("Input").Cells(41, 1).Value)
                If i = vbCancel Then Exit Sub
                If i = vbYes Then
                    Worksheets("Input").Cells(6, 2).Value = "Yes"
                Else
                    Worksheets("Input").Cells(6, 2).Value = "No"
                End If
                Worksheets("Input").Cells(41, 1).Value = i
                If i = vbYes Then 'adjust
                    Worksheets("Storage").Cells(1, 1).Value = 0
                Else 'use the input grain size anyway
                    Worksheets("Storage").Cells(1, 1).Value = 1
                End If
            End If
            Worksheets("Storage").Columns("C:D").ClearContents
            If Worksheets("Storage").Cells(1, 1).Value = 0 Then 'adjust grain size
                If Worksheets("Input").Cells(1, 5).Value > Worksheets("Input").Cells(2, 5).Value Then
                    'degreasing grain size
                    k = 0
                    Do While Worksheets("Input").Cells(k + 1, 5).Value >= 2
                        k = k + 1
                    Loop
                    For i = 1 To k
                        Worksheets("Storage").Cells(i, 3).Value = Worksheets("Input").Cells(i, 5).Value
                        Worksheets("Storage").Cells(i, 4).Value = _
                            100 - (100 - Worksheets("input").Cells(i, 6).Value) / _
                            (100 - Worksheets("input").Cells(k, 6).Value) * 100
                    Next
                Else
                    'increasing grain size
                    k = 1
                    Do While Worksheets("Input").Cells(k, 5).Value < 2
                        k = k + 1
                    Loop
                    i = k - 1
                    Do While Not IsEmpty(Worksheets("Input").Cells(i + 1, 5))
                        i = i + 1
                        Worksheets("Storage").Cells(i - k + 1, 3).Value = Worksheets("Input").Cells(i, 5).Value
                        Worksheets("Storage").Cells(i - k + 1, 4).Value = _
                            100 - (100 - Worksheets("Input").Cells(i, 6).Value) / _
                            (100 - Worksheets("Input").Cells(k, 6).Value) * 100
                    Loop
                End If
            Else ' no adjust is needed
                i = 0
                Do While Not IsEmpty(Worksheets("Input").Cells(i + 1, 5))
                    i = i + 1
                    Worksheets("storage").Cells(i, 3).Value = Worksheets("Input").Cells(i, 5).Value
                    Worksheets("storage").Cells(i, 4).Value = Worksheets("Input").Cells(i, 6).Value
                Loop
            End If
        End If
        
        CalculateSurfaceD65AndGravelSandFractions 1
        
        CheckOutSubstrateGrainSizeNeed 1
        Exit Sub
    End If

    If Cells(1, 1).Value = "Size (mm)" And Left(Cells(3, 4).Value, 3) = "Sub" Then
    'substrate grain size distribution
        i = 1
        Do While Not IsEmpty(Cells(i + 1, 1))
            i = i + 1
            Worksheets("Input").Cells(i - 1, 7).Value = Cells(i, 1).Value
            Worksheets("Input").Cells(i - 1, 8).Value = Cells(i, 2).Value
        Loop
        Range(Worksheets("Input").Cells(i, 7), _
            Worksheets("Input").Cells(Worksheets("Input").Rows.Count, 8)).ClearContents
        ShowAndHide Worksheets("Welcome"), Worksheets("MyInput")
        
        'A portion of code for adjust substrate grain size to exclude sand is
        'deleted from here and saved in a txt file in case needed later.

        CalculateSubstrateMeanGrainSize 1
        
        GoGetBakkesSamplingInformation 1
        Exit Sub
    End If
    
    If Left(Cells(1, 1).Value, 5) = "Disch" Then 'discharge
        If Right(Cells(1, 1).Value, 5) = "(cms)" Then
            QwUnit = 1
        Else ' in cfs
            QwUnit = 0.3048 ^ 3
        End If
        If Worksheets("Input").Cells(6, 1).Value = "(C1)" Then ' duration curve
            j = 1
            Do While Not IsEmpty(Cells(j + 1, 1))
                j = j + 1
            Loop
            If Abs(Abs(Cells(2, 2).Value - Cells(j, 2).Value) - 100) > 0.00001 Then
                MsgBox "Error detected!  Exceedance probabilities must be bounded by 0 at one end and " & _
                    "100 at the other.", vbOKOnly + vbCritical, "Discharge"
                Exit Sub
            End If
            If (Cells(2, 2).Value > Cells(j, 2).Value And Cells(2, 1).Value > Cells(j, 1).Value) Or _
                (Cells(2, 2).Value < Cells(j, 2).Value And Cells(2, 1).Value < Cells(j, 1).Value) Then
                MsgBox "Error detected!  Maximum discharge should be associated with 0 " & _
                    "exceedance probability and minimum discharge should be associated " & _
                    "with 100% exceedance probability.", vbOKOnly + vbCritical, "Discharge"
                Exit Sub
            End If
            For i = 1 To 26
                For k = 2 To j - 1
                    If (Cells(k, 2).Value <= Worksheets("Input").Cells(i, 17).Value And _
                        Cells(k + 1, 2).Value >= Worksheets("Input").Cells(i, 17).Value) Or _
                        (Cells(k, 2).Value >= Worksheets("Input").Cells(i, 17).Value And _
                        Cells(k + 1, 2).Value <= Worksheets("Input").Cells(i, 17).Value) Then
                        Worksheets("Input").Cells(i, 16).Value = _
                            (Cells(k, 1).Value + (Cells(k + 1, 1).Value - Cells(k, 1).Value) / _
                            (Cells(k + 1, 2).Value - Cells(k, 2).Value) * _
                            (Worksheets("Input").Cells(i, 17).Value - Cells(k, 2).Value)) * QwUnit
                        Exit For
                    End If
                Next
            Next
        Else 'discharge record is provided
            Application.ScreenUpdating = False
            j = 1
            Do While Not IsEmpty(Cells(j + 1, 1))
                j = j + 1
            Loop ' j is the last row number
            If j < 365 Then
                MsgBox "You have less than a year's discharge record.  The calculated " & _
                    "average bedload transport rate can only represent this specific " & _
                    "period and should not be extrapolated to other time periods.", _
                    vbOKOnly + vbInformation, "Discharge"
            End If
            UnprotectThisSheet 1
            Range(Cells(2, 1), Cells(j, 1)).Select
            Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlNo, _
                OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
            Worksheets("Input").Cells(1, 16).Value = Cells(2, 1).Value * QwUnit
            Worksheets("Input").Cells(26, 16).Value = Cells(j, 1).Value * QwUnit
            For i = 2 To 25
                dmy = (1 - Worksheets("Input").Cells(i, 17).Value / 100) * j
                k = Int(dmy) 'record no.
                k = k + 1 'row number
                If k < 2 Then
                    Worksheets("Input").Cells(i, 16).Value = Cells(2, 1).Value * QwUnit
                ElseIf k >= j Then
                    Worksheets("Input").Cells(i, 16).Value = Cells(j, 1).Value * QwUnit
                Else
                    Worksheets("Input").Cells(i, 16).Value = _
                        (Cells(k, 1).Value + (Cells(k, 1).Value - Cells(k, 1).Value) * _
                        (dmy + 1 - k)) * QwUnit
                End If
            Next
            Range(Cells(2, 1), Cells(Rows.Count, 1)).ClearContents
            Application.ScreenUpdating = True
        End If
        
        ShowAndHide Worksheets("Welcome"), Worksheets("MyInput")
        
        ManipulateCrossSection

    End If

End Sub

Sub CancelClickedOnMyInput()
    If EnableOnError Then On Error Resume Next
    ShowAndHide Worksheets("Welcome"), Worksheets("MyInput")
    UserFormInUse = False
    Canceled = True
End Sub

Sub ClearFormClickedOnMyInput()
    If EnableOnError Then On Error Resume Next
    UnprotectThisSheet 1
    Range(Cells(2, 1), Cells(ActiveSheet.Rows.Count, 2)).ClearContents
    ActiveSheet.ScrollArea = "A2:B" & Rows.Count
    Columns("B:B").Hidden = False
    Cells(2, 1).Select
    ProtectThisSheet 1
End Sub

Private Sub ShowCrossSectionAndGetFloodplainInformation()
    Dim i As Long, MinX As Long, MinY As Long, MaxX As Long, MaxY As Long
    Dim MyRange As Range
    
    If EnableOnError Then On Error Resume Next
    Application.ScreenUpdating = False
    
    ShowAndHide Worksheets("PlotXS"), Worksheets("Welcome")
    UnprotectThisSheet 1
    
    Columns("J:K").ClearContents
    
    i = 1
    Do While Not IsEmpty(Worksheets("Input").Cells(i + 1, 3))
        i = i + 1
    Loop
    
    Set MyRange = Range(Worksheets("Input").Cells(1, 3), Worksheets("Input").Cells(i, 3))
    MinX = Int(Application.WorksheetFunction.Min(MyRange) / 10 - 1) * 10
    MaxX = Int(Application.WorksheetFunction.Max(MyRange) / 10 + 1) * 10
    
    Set MyRange = Range(Worksheets("Input").Cells(1, 4), Worksheets("Input").Cells(i, 4))
    MinY = Int(Application.WorksheetFunction.Min(MyRange) - 1)
    MaxY = Int(Application.WorksheetFunction.Max(MyRange) + 1)
    
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.SeriesCollection(1).Formula = _
        "=SERIES(,Input!R1C3:R" & i & "C3,Input!R1C4:R" & i & "C4,1)"
    With ActiveChart.Axes(xlValue)
        .MinimumScale = MinY
        .MaximumScale = MaxY
        .CrossesAt = MinY
    End With
    With ActiveChart.Axes(xlCategory)
        .MinimumScale = MinX
        .MaximumScale = MaxX
        .CrossesAt = MinX
    End With
    ProtectThisSheet 1
    ActiveChart.Refresh
    Application.ScreenUpdating = True
    
    ActiveSheet.ScrollArea = "A22"
    Worksheets("PlotXS").Cells(23, 1).Value = "FALSE"
    Worksheets("PlotXS").Cells(24, 1).Value = "FALSE"
    Worksheets("PlotXS").Range("F23:G24").ClearContents
    ActiveSheet.Cells(22, 1).Select
    
    i = MyMsgBox("Here is the cross section you just entered.  " & vbLf & vbLf & _
        "Click Yes if your cross section includes floodplain(s), " & _
        "and enter roughness and distance information below.  Click No " & _
        "if the cross section does not have a floodplain.", "BAGS", _
        Worksheets("Input").Cells(41, 2).Value)
    If i = vbCancel Then
        Auto_Open
        Exit Sub
    End If
    Worksheets("Input").Cells(41, 2).Value = i
    If i = vbNo Then
        Worksheets("Input").Cells(4, 1).Value = Application.WorksheetFunction.Min( _
            Range(Worksheets("Input").Cells(1, 3), Worksheets("Input").Cells(Rows.Count, 3)))
        Worksheets("Input").Cells(4, 2).Value = Application.WorksheetFunction.Max( _
            Range(Worksheets("Input").Cells(1, 3), Worksheets("Input").Cells(Rows.Count, 3)))
        If IsEmpty(Worksheets("Input").Cells(3, 1)) Then
            Worksheets("Input").Cells(3, 1).Value = 0.07
        End If
        If IsEmpty(Worksheets("Input").Cells(3, 2)) Then
            Worksheets("Input").Cells(3, 2).Value = 0.07
        End If
        ShowAndHide Worksheets("Welcome"), Worksheets("PlotXS")
        GoGetGrainSizeInformation 1
    Else
        PlotFloodplains 1
        If Worksheets("Input").Cells(2, 1).Value = "Yes" Then
            Worksheets("PlotXS").Cells(23, 1).Value = "TRUE"
            Worksheets("PlotXS").Cells(23, 6).Value = Worksheets("Input").Cells(3, 1).Value
            Worksheets("PlotXS").Cells(23, 7).Value = Worksheets("Input").Cells(4, 1).Value
        End If
        If Worksheets("Input").Cells(2, 2).Value = "Yes" Then
            Worksheets("PlotXS").Cells(24, 1).Value = "TRUE"
            Worksheets("PlotXS").Cells(24, 6).Value = Worksheets("Input").Cells(3, 2).Value
            Worksheets("PlotXS").Cells(24, 7).Value = Worksheets("Input").Cells(4, 2).Value
        End If
        If Worksheets("Input").Cells(2, 1).Value = "Yes" And Worksheets("Input").Cells(2, 1).Value = "Yes" Then
            Worksheets("PlotXS").ScrollArea = "F23:G24"
        End If
        If Worksheets("Input").Cells(2, 1).Value = "Yes" And Worksheets("Input").Cells(2, 1).Value = "No" Then
            Worksheets("PlotXS").ScrollArea = "F23:G23"
        End If
        If Worksheets("Input").Cells(2, 1).Value = "No" And Worksheets("Input").Cells(2, 1).Value = "Yes" Then
            Worksheets("PlotXS").ScrollArea = "F24:G24"
        End If
    End If
    UnprotectThisSheet 1
    If ActiveSheet.Name = "PlotXS" Then ActiveCell.Interior.ColorIndex = 15
    ProtectThisSheet 1
End Sub

Sub ContinueOnPlotXSClicked()
    Dim i As Long
    If EnableOnError Then On Error Resume Next
    If Cells(23, 1) Then
        Worksheets("Input").Cells(2, 1).Value = "Yes"
        Worksheets("Input").Cells(3, 1).Value = Cells(23, 6).Value
        Worksheets("Input").Cells(4, 1).Value = Cells(23, 7).Value
    Else
        Worksheets("Input").Cells(2, 1).Value = "No"
        Worksheets("Input").Cells(3, 1).Value = 0.07
        Worksheets("Input").Cells(4, 1).Value = Application.WorksheetFunction.Min( _
            Range(Worksheets("Input").Cells(1, 3), Worksheets("Input").Cells(Rows.Count, 3)))
    End If
    If Cells(24, 1) Then
        Worksheets("Input").Cells(2, 2).Value = "Yes"
        Worksheets("Input").Cells(3, 2).Value = Cells(24, 6).Value
        Worksheets("Input").Cells(4, 2).Value = Cells(24, 7).Value
    Else
        Worksheets("Input").Cells(2, 2).Value = "No"
        Worksheets("Input").Cells(3, 2).Value = 0.07
        Worksheets("Input").Cells(4, 2).Value = Application.WorksheetFunction.Max( _
            Range(Worksheets("Input").Cells(1, 3), Worksheets("Input").Cells(Rows.Count, 3)))
    End If
    If Not (Cells(23, 1).Value Or Cells(24, 1).Value) Then
        i = MyMsgBox("Are you sure you do not have floodplains in this cross section?", _
            "BAGS", vbNo)
        If i = vbNo Then Exit Sub
    End If
    ShowAndHide Worksheets("Welcome"), Worksheets("PlotXS")
    GoGetGrainSizeInformation 1
End Sub

Sub CancelOnPlotXSClicked()
    If EnableOnError Then On Error Resume Next
    ShowAndHide Worksheets("Welcome"), Worksheets("PlotXS")
End Sub

Sub HelpOnPlotXSClicked()
    If EnableOnError Then On Error Resume Next
    Load ufManning
    ufManning.Show
End Sub

Sub cbOnPlotXSClicked()
    If EnableOnError Then On Error Resume Next
    UnprotectThisSheet 1
    Range(Cells(23, 6), Cells(24, 7)).Interior.ColorIndex = 2
    If Cells(23, 1).Value And Cells(24, 1).Value Then
        ActiveSheet.ScrollArea = "F23:G24"
        Cells(23, 6).Select
        Cells(23, 6).Interior.ColorIndex = 15
    ElseIf Cells(23, 1).Value And Not Cells(24, 1).Value Then
        Cells(24, 6).ClearContents: Cells(24, 7).ClearContents
        Range(Cells(24, 6), Cells(24, 7)).Interior.ColorIndex = 15
        ActiveSheet.ScrollArea = "F23:G23"
        Cells(23, 6).Select
        Cells(23, 6).Interior.ColorIndex = 15
    ElseIf Not Cells(23, 1).Value And Cells(24, 1).Value Then
        Cells(23, 6).ClearContents: Cells(23, 7).ClearContents
        Range(Cells(23, 6), Cells(23, 7)).Interior.ColorIndex = 15
        ActiveSheet.ScrollArea = "F24:G24"
        Cells(24, 6).Select
        Cells(24, 6).Interior.ColorIndex = 15
    Else
        Range(Cells(23, 6), Cells(24, 7)).ClearContents
        Range(Cells(23, 6), Cells(24, 7)).Interior.ColorIndex = 15
        ActiveSheet.ScrollArea = "A1"
    End If
    ProtectThisSheet 1
    PlotFloodplains 1
End Sub

Sub GoGetGrainSizeInformation(dmy As Integer)
    If EnableOnError Then On Error Resume Next
    Dim i As Long
    
    On Error Resume Next
    
    If Parker90 Or Wilcock03 Then
        GoGetSurfaceGrainSizeDistribution 1
    ElseIf Wilcock Then
        i = MyMsgBox("To use the two-fraction equation of Wilcock (2001), you need " & _
            "to supply gravel and sand fractions on channel bed, and an estimate " & _
            "of surface grain size D65.  You can also supply a full " & _
            "grain size distribution of the channel surface." & vbLf & vbLf & _
            "Supply gravel/sand fractions and an estimate of D65? (Click ""No"" to enter a full surface grain " & _
            "size distritution or ""Cancel"" to stop the run.)", _
            "Wilcock (2001)", Worksheets("Input").Cells(42, 1).Value)
        If i = vbCancel Then Exit Sub
        Worksheets("Input").Cells(42, 1).Value = i
        If i = vbNo Then 'use full surface grain size distribution
            GoGetSurfaceGrainSizeDistribution 1
        ElseIf i = vbYes Then 'use gravel/sand fractions
            UserFormInUse = True
            Canceled = False
            Load ufFractions
            ufFractions.tbD65.Value = Format(Worksheets("Input").Cells(13, 2).Value, "##0.0")
            ufFractions.tbGravel.Value = Format(Worksheets("Input").Cells(12, 1).Value, "0.##0")
            ufFractions.tbSand.Value = Format(1 - ufFractions.tbGravel.Value, "0.##0")
            ufFractions.Show
            
            Do While UserFormInUse
                DoEvents
            Loop
            
            If Canceled Then Exit Sub
            
            Worksheets("Input").Columns("E:F").ClearContents
            
            CheckOutSubstrateGrainSizeNeed 1
        End If
    Else ' substrate based equations
        CheckOutSubstrateGrainSizeNeed 1
    End If
End Sub

Sub GoGetSurfaceGrainSizeDistribution(dmy As Integer)
    Dim i As Long
    If EnableOnError Then On Error Resume Next
    Application.ScreenUpdating = False
    ShowAndHide Worksheets("MyInput"), Worksheets("Welcome")
    ClearFormClickedOnMyInput
    Cells(1, 1).Value = "Size (mm)"
    Cells(1, 2).Value = "% Finer"
    Cells(3, 4).Value = "Surface Grain Size Distribution" & vbLf & vbLf & _
        "Based on the equation(s) you selected, you need to " & _
        "provide surface layer grain size distribution in order to carry on " & _
        "the calculation." & vbLf & vbLf & _
        "It is suggested that grain sizes be provided in a half-phi interval " & _
        "to a one-phi interval.  Grain sizes in a half-phi interval, for example, " & _
        "would be ..., 2, 2.8, 4, 5.6, 8, 11, 16 mm, ..., and grain size in a " & _
        "one-phi interval would be ..., 2, 4, 8, 16, 32, 64 mm, ..." & vbLf & vbLf & _
        "The percent finers associated to the finest and coarsest grain sizes " & _
        "must be 0 and 100, respectively." & vbLf & vbLf & _
        "Click ""Accept"" to continue upon finishing the input, or click " & _
        """Cancel"" to quit the calculation."
    ActiveSheet.ScrollArea = "A2:B" & ActiveSheet.Rows.Count
    If Not IsEmpty(Worksheets("Input").Cells(1, 5)) Then
        i = 0
        Do While Not IsEmpty(Worksheets("Input").Cells(i + 1, 5))
            i = i + 1
            Cells(i + 1, 1).Value = Worksheets("Input").Cells(i, 5).Value
            Cells(i + 1, 2).Value = Worksheets("Input").Cells(i, 6).Value
        Loop
    End If
    Application.ScreenUpdating = True
End Sub

Sub CheckOutSubstrateGrainSizeNeed(dmy As Integer)
    If EnableOnError Then On Error Resume Next
    If PK82 Or Bakke Then
        GoGetFullSubstrateGrainSizeDistribution 1
    Else
        If Parker82 Then
            GoGetSubstrateGrainSizeDistribution 1
        Else 'skip substrate
            If Bakke Or Wilcock Then
                GoGetBakkesSamplingInformation 1
            Else
                GoGetSlopeAndDischarge 1
            End If
        End If
    End If
End Sub

Sub GoGetSubstrateGrainSizeDistribution(dmy As Integer)
    Dim i As Long, D50
    If EnableOnError Then On Error Resume Next
    i = MyMsgBox("To apply Parker-Klingeman-McLean (1982) equation or Bakke (1999) procedure, " & _
        "you must supply a median grain size (D50) for substrate.  Or alternatively you can supply " & _
        "a full substrate grain size distribution." & vbLf & vbLf & _
        "Use median grain size? (Click ""No"" to use full grain size distribution and " & _
        """Cancel"" to stop the run.)", "Substrate size", Worksheets("Input").Cells(42, 2).Value)
    If i = vbCancel Then Exit Sub
    Worksheets("Input").Cells(42, 2).Value = 1
    If i = vbNo Then 'use full size distribution
        GoGetFullSubstrateGrainSizeDistribution 1
    ElseIf i = vbYes Then 'use D50
        D50 = Application.InputBox("Enter substrate median grain size (D50), in mm:", _
            "Substrate median size", Worksheets("Input").Cells(13, 1).Value)
        If UCase(D50) = 0 Then Exit Sub
        Worksheets("Input").Cells(13, 1).Value = D50
        Worksheets("Input").Columns("G:H").ClearContents
        GoGetBakkesSamplingInformation 1
    End If
End Sub

Sub GoGetFullSubstrateGrainSizeDistribution(dmy As Integer)
    Dim i As Long
    If EnableOnError Then On Error Resume Next
    Application.ScreenUpdating = False
    ShowAndHide Worksheets("MyInput"), Worksheets("Welcome")
    ClearFormClickedOnMyInput
    Cells(1, 1).Value = "Size (mm)"
    Cells(1, 2).Value = "% Finer"
    Cells(3, 4).Value = "Substrate Grain Size Distribution" & vbLf & vbLf & _
        "Based on the equation(s) you selected, you need to " & _
        "provide substrate grain size distribution in order to carry on the " & _
        "calculation." & vbLf & vbLf & _
        "It is suggested that grain sizes be provided in a half-phi interval " & _
        "to a one-phi interval.  Grain sizes in a half-phi interval, for example, " & _
        "would be ..., 2, 2.8, 4, 5.6, 8, 11, 16 mm, ..., and grain size in a " & _
        "one-phi interval would be ..., 2, 4, 8, 16, 32, 64 mm, ..." & vbLf & vbLf & _
        "The percent finers associated to the finest and coarsest grain sizes " & _
        "must be 0 and 100, respectively." & vbLf & vbLf & _
        "Click ""Accept"" to continue upon finishing the input, or click " & _
        """Cancel"" to quit the calculation."
    ActiveSheet.ScrollArea = "A2:B" & ActiveSheet.Rows.Count
    If Not IsEmpty(Worksheets("Input").Cells(1, 7)) Then
        i = 0
        Do While Not IsEmpty(Worksheets("Input").Cells(i + 1, 7))
            i = i + 1
            Cells(i + 1, 1).Value = Worksheets("Input").Cells(i, 7).Value
            Cells(i + 1, 2).Value = Worksheets("Input").Cells(i, 8).Value
        Loop
    End If
    Application.ScreenUpdating = True
End Sub

Sub GoGetBakkesSamplingInformation(dmy As Integer)
    Dim i As Long, Nsize As Long
    Dim NSamples As Long
    Dim InSt As Worksheet, MyRange As Range
    
    If EnableOnError Then On Error Resume Next
    Set InSt = Worksheets("Input")

    If Bakke Then
        Nsize = 0
        Do While Not IsEmpty(InSt.Cells(Nsize + 1, 7))
            Nsize = Nsize + 1
        Loop
        Worksheets("Bedload").Cells(1, 2).Value = Nsize
    End If
    
    If Bakke Or Wilcock Then
        NSamples = -9991
        Do While NSamples < 2
            If NSamples <> -9991 Then _
                MsgBox "There must be at least 2 sampling results!", _
                    vbOKOnly + vbInformation, "Bedload sampleing"
            NSamples = InSt.Cells(7, 1).Value
            NSamples = Application.InputBox("Enter the number of bedload transport samples:", _
                "Number of samples", NSamples)
            If NSamples = 0 Then Exit Sub
            If NSamples >= 2 Then Worksheets("Input").Cells(7, 1).Value = NSamples
        Loop
        MessageOnWelcome "Loading data!  Please wait..."
        Cells(1, 1).Select
        Application.ScreenUpdating = False
        ShowAndHide Worksheets("Bedload"), Worksheets("Welcome")
        UnprotectThisSheet 1
        
        ActiveSheet.Rows.Hidden = False
        Cells.Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        Selection.Borders(xlEdgeLeft).LineStyle = xlNone
        Selection.Borders(xlEdgeTop).LineStyle = xlNone
        Selection.Borders(xlEdgeBottom).LineStyle = xlNone
        Selection.Borders(xlEdgeRight).LineStyle = xlNone
        Selection.Borders(xlInsideVertical).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        Selection.Interior.ColorIndex = 15
        
        
        If Bakke Then
            ActiveSheet.ScrollArea = "E14:E" & 19 + Nsize
            ActiveSheet.Rows("16:17").Hidden = True
            Set MyRange = Application.Union( _
                Range(Cells(14, 5), Cells(15, 5)), _
                Range(Cells(20, 5), Cells(Nsize + 19, 5)))
            MyRange.Select
        Else
            ActiveSheet.ScrollArea = "E14:E16"
            ActiveSheet.Rows("19:" & Rows.Count).Hidden = True
            Range(Cells(14, 5), Cells(16, 5)).Select
        End If
        Selection.Interior.ColorIndex = 2
        
        If Bakke Then
            Range(Cells(20, 4), Cells(100, 5)).ClearContents
            For i = 1 To Nsize
                Cells(19 + i, 4).Value = InSt.Cells(i, 7).Value
            Next
            Set MyRange = Application.Union( _
                Range(Cells(14, 5), Cells(15, 5)), _
                Range(Cells(19, 4), Cells(19 + Nsize, 5)))
            MyRange.Select
        Else
            Range(Cells(14, 5), Cells(17, 5)).Select
        End If
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
            
        Range(Cells(14, 5), Cells(16, 5)).ClearContents
        
        Cells(1, 1).Value = 1
        Cells(12, 2).Value = "Bedload sampling record No. 1"
        
        Cells(14, 5).Select
        ActiveSheet.Shapes("Button 1").Visible = False
        If Cells(1, 1).Value = Worksheets("Input").Cells(7, 1).Value Then
            ActiveSheet.Shapes("FinishButton").Visible = True
            ActiveSheet.Shapes("NextButton").Visible = False
        Else
            ActiveSheet.Shapes("FinishButton").Visible = False
            ActiveSheet.Shapes("NextButton").Visible = True
        End If
        ProtectThisSheet 1
        Range(InSt.Cells(NSamples + 1, 18), InSt.Cells(Rows.Count, 40)).ClearContents
        Range(InSt.Cells(1, 21 + Nsize), InSt.Cells(Rows.Count, 39)).ClearContents
        
        RetriveBakkeSampling Cells(1, 1)
        Application.ScreenUpdating = True
    Else
        GoGetSlopeAndDischarge 1
    End If
End Sub

Sub PreviousClickedOnBedload()
    Dim i As Long, dmy As Boolean
    
    If EnableOnError Then On Error Resume Next
    dmy = SaveBakkeSampling(Cells(1, 1).Value)
    
    ClearFormClickedOnBedload
    ActiveSheet.Shapes("Button 1").Visible = True
    i = Cells(1, 1).Value
    Cells(1, 1).Value = i - 1
    Cells(12, 2).Value = "Bedload sampling record No. " & Cells(1, 1).Value
    If Cells(1, 1).Value = 1 Then ActiveSheet.Shapes("Button 1").Visible = False
    
    If Cells(1, 1).Value = Worksheets("Input").Cells(7, 1).Value Then
        ActiveSheet.Shapes("FinishButton").Visible = True
        ActiveSheet.Shapes("NextButton").Visible = False
    Else
        ActiveSheet.Shapes("FinishButton").Visible = False
        ActiveSheet.Shapes("NextButton").Visible = True
    End If
    
    RetriveBakkeSampling Cells(1, 1)
End Sub

Sub NextClickedOnBedload()
    Dim i As Long
    
    If EnableOnError Then On Error Resume Next
    If Not SaveBakkeSampling(Cells(1, 1).Value) Then
        MsgBox "Incomplete input!  Although sampling grain size distribution is " & _
        "optional, both discharge and bedload transport rate must be supplied " & _
            "before you can proceed to the next record.", vbOKOnly + vbCritical, _
            "Incomplete input"
        Exit Sub
    End If
    
    ClearFormClickedOnBedload
    ActiveSheet.Shapes("Button 1").Visible = True
    i = Cells(1, 1).Value
    Cells(1, 1).Value = i + 1
    Cells(12, 2).Value = "Bedload sampling record No. " & Cells(1, 1).Value
    
    If Cells(1, 1).Value = Worksheets("Input").Cells(7, 1).Value Then
        ActiveSheet.Shapes("FinishButton").Visible = True
        ActiveSheet.Shapes("NextButton").Visible = False
    Else
        ActiveSheet.Shapes("FinishButton").Visible = False
        ActiveSheet.Shapes("NextButton").Visible = True
    End If
    
    RetriveBakkeSampling Cells(1, 1).Value
    
End Sub

Function SaveBakkeSampling(RecNo As Long) As Boolean
    'assume "Bedload" is active sheet
    Dim pG As Double, Nsize As Long, InSt As Worksheet, i As Long
    
    On Error Resume Next
    Set InSt = Worksheets("Input")
    
    If IsEmpty(Cells(14, 5)) Or IsEmpty(Cells(15, 5)) Then
        SaveBakkeSampling = False
        Exit Function
    End If
        
    If Bakke Then 'interpolate for gravel/sand fraction
        Nsize = Cells(1, 2).Value
        CalculateSubstrateGravelSandFraction Nsize, pG
        Cells(16, 5).Value = pG
    End If
    
    If Wilcock And IsEmpty(Cells(16, 5)) Then
        SaveBakkeSampling = False
        Exit Function
    End If

    InSt.Cells(RecNo, 18).Value = Cells(14, 5).Value
    InSt.Cells(RecNo, 19).Value = Cells(15, 5).Value
    InSt.Cells(RecNo, 20).Value = Cells(16, 5).Value
    If Bakke Then
        For i = 1 To Nsize
            InSt.Cells(RecNo, 20 + i).Value = Cells(19 + i, 5).Value
        Next
    End If
    InSt.Cells(RecNo, 40).Value = Cells(2, 1).Value
    SaveBakkeSampling = True
End Function

Sub RetriveBakkeSampling(RecNo As Long) 'assume "Bedload" is active sheet
    Dim i As Long, Nsize As Long
    If EnableOnError Then On Error Resume Next
    Nsize = Cells(1, 2).Value
    If IsEmpty(Worksheets("Input").Cells(RecNo, 40)) Then
        Cells(2, 1).Value = False
    Else
        Cells(2, 1).Value = Worksheets("Input").Cells(RecNo, 40).Value
    End If
    Cells(14, 5).Value = Worksheets("Input").Cells(RecNo, 18).Value
    Cells(15, 5).Value = Worksheets("Input").Cells(RecNo, 19).Value
    If Bakke Then
        For i = 1 To Nsize
            Cells(19 + i, 5).Value = Worksheets("Input").Cells(RecNo, 20 + i).Value
        Next
    End If
    If Wilcock Then _
        Cells(16, 5).Value = Worksheets("Input").Cells(RecNo, 20).Value
End Sub

Sub ClearFormClickedOnBedload()
    Dim Nsize As Long
    If EnableOnError Then On Error Resume Next
    Nsize = Cells(1, 2).Value
    Application.ScreenUpdating = False
    Range(Cells(14, 5), Cells(15, 5)).ClearContents
    Range(Cells(21, 5), Cells(20 + Nsize, 5)).ClearContents
    Cells(16, 5).Value = 1
    Cells(14, 5).Select
    Application.ScreenUpdating = True
End Sub

Sub FinishClickedOnBedload()
    If EnableOnError Then On Error Resume Next
    If Not SaveBakkeSampling(Cells(1, 1).Value) Then
        MsgBox "Incomplete input!  Although sampling grain size distribution is " & _
        "optional, both discharge and bedload transport rate must be supplied " & _
            "before you can proceed to finish the input for sampling record.", vbOKOnly + vbCritical, _
            "Incomplete input"
        Exit Sub
    End If
    ShowAndHide Worksheets("Welcome"), Worksheets("Bedload")
    GoGetSlopeAndDischarge 1
End Sub

Sub CancelClickedOnBedload()
    If EnableOnError Then On Error Resume Next
    ShowAndHide Worksheets("Welcome"), Worksheets("Bedload")
End Sub

Sub GoGetSlopeAndDischarge(dmy As Integer)
    Dim i As Long
    If EnableOnError Then On Error Resume Next
    UserFormInUse = True
    Canceled = False
    Load ufSlope
    ufSlope.tbSlope = Worksheets("Input").Cells(5, 2).Value
    If Worksheets("Input").Cells(5, 1).Value = "W.S." Then ufSlope.obWS.Value = True
    If Worksheets("Input").Cells(5, 1).Value = "Bed" Then ufSlope.obBed.Value = True
    If Worksheets("Input").Cells(5, 1).Value = "Model" Then ufSlope.obModel.Value = True
    ufSlope.Show
    Do While UserFormInUse
        DoEvents
    Loop
    If Canceled Then Exit Sub
    
    Cells(1, 1).Select
    UserFormInUse = True
    Canceled = False
    Load ufDischarge
    If Worksheets("Input").Cells(6, 1).Value = "(A)" Then
        ufDischarge.obA.Value = True
        ufDischarge.Frame1.Visible = False
    ElseIf Worksheets("Input").Cells(6, 1).Value = "(B)" Then
        ufDischarge.obB.Value = True
        ufDischarge.Frame1.Visible = False
    Else '(C1) or (C2)
        ufDischarge.obC.Value = True
        ufDischarge.Frame1.Visible = True
        ufDischarge.obDuration.Value = True
    End If
    ufDischarge.Show
    Do While UserFormInUse
        DoEvents
    Loop
    
    If Canceled Then Exit Sub
    
    Cells(1, 1).Select
    UserFormInUse = True
    Canceled = False
    If Worksheets("Input").Cells(6, 1).Value = "(A)" Then
        Load ufSingleDischarge
        ufSingleDischarge.obCMS.Value = True
        ufSingleDischarge.lbUnit.Caption = "cms"
        If Not IsEmpty(Worksheets("Input").Cells(1, 16)) Then _
            ufSingleDischarge.tbQw.Value = _
                Format(Worksheets("Input").Cells(1, 16).Value, "###0.##")
        ufSingleDischarge.Show
    
        Do While UserFormInUse
            DoEvents
        Loop
        If Canceled Then Exit Sub
        
        ManipulateCrossSection
        
    ElseIf Worksheets("Input").Cells(6, 1).Value = "(B)" Then
        Load ufMinMaxDischarge
        ufMinMaxDischarge.obCMS.Value = True
        ufMinMaxDischarge.lbUnit1.Caption = "cms"
        ufMinMaxDischarge.lbUnit2.Caption = "cms"
        If Not IsEmpty(Worksheets("Input").Cells(1, 16)) Then _
            ufMinMaxDischarge.tbMinQw.Value = _
            Format(Worksheets("Input").Cells(1, 16).Value, "###0.##")
        If Not IsEmpty(Worksheets("Input").Cells(26, 16)) Then _
            ufMinMaxDischarge.tbMaxQw.Value = _
            Format(Worksheets("Input").Cells(26, 16).Value, "###0.##")
        ufMinMaxDischarge.Show
        
        Do While UserFormInUse
            DoEvents
        Loop
        If Canceled Then Exit Sub
        
        ManipulateCrossSection
        
    ElseIf Worksheets("Input").Cells(6, 1).Value = "(C1)" Then
        UserFormInUse = False
        Application.ScreenUpdating = False
        ShowAndHide Worksheets("MyInput"), Worksheets("Welcome")
        UnprotectThisSheet 1
        ActiveSheet.Shapes("obCMS").Visible = True
        ActiveSheet.Shapes("obCFS").Visible = True
        Cells(1, 10).Value = 1 '1 for cms and 2 for cfs
        Cells(1, 1).Value = "Discharge (cms)"
        Cells(1, 2).Value = "Exceedance Probability (%)"
        Cells(3, 4).Value = "Please select unit for discharge" & vbLf & _
            "BEFORE entering discharge data:" & vbLf & vbLf & _
            "The suggested exceedance probabilities for input are given on " & _
            "the left.  You are, however, allowed to use different exceedance " & _
            "probabilities.  The software will interpolate your input to the suggested " & _
            "exceedance probabilities if different values are supplied." & _
            vbLf & vbLf & "If discharge data are loaded, please check if " & _
            "they are the data you intended to supply because the last run may " & _
            "have supplied discharge data from an unrelated project." & vbLf & vbLf & _
            "Click ""Accept"" to continue upon finishing your input."
        ClearFormClickedOnMyInput
        ProtectThisSheet 1
        If Not IsEmpty(Worksheets("Input").Cells(3, 16)) Then 'load data
            i = 0
            Do While Not IsEmpty(Worksheets("Input").Cells(i + 1, 16))
                i = i + 1
                Cells(i + 1, 1).Value = Worksheets("Input").Cells(i, 16).Value
                Cells(i + 1, 2).Value = Worksheets("Input").Cells(i, 17).Value
            Loop
        Else
            For i = 1 To 26
                Cells(i + 1, 2).Value = Worksheets("Input").Cells(i, 17).Value
            Next
        End If
        Application.ScreenUpdating = True
    ElseIf Worksheets("Input").Cells(6, 1).Value = "(C2)" Then
        UserFormInUse = False
        Application.ScreenUpdating = False
        ShowAndHide Worksheets("MyInput"), Worksheets("Welcome")
        ClearFormClickedOnMyInput
        UnprotectThisSheet 1
        Columns("B:B").Hidden = True
        ActiveSheet.Shapes("obCMS").Visible = True
        ActiveSheet.Shapes("obCFS").Visible = True
        Cells(1, 10).Value = 1 '1 for cms and 2 for cfs
        Cells(1, 1).Value = "Discharge (cms)"
        Cells(3, 4).Value = "Please select unit for discharge" & vbLf & _
            "BEFORE entering discharge data." & vbLf & vbLf & _
            "Upon selecting appropriate unit for discharge, enter or copy discharge " & _
            "into the space provided (column A, starting row no. 2).  Click ""Accept"" " & _
            "upon finishing." & vbLf & vbLf & _
            "Please be adviced that the software will transfer your discharge " & _
            "record into a flow duration curve upon clicking ""Accept"", and the " & _
            "discharge record will not be saved in this software.  If you make " & _
            "mistakes somwhere along the way and restart the run, you can use the " & _
            "flow duration curve generated with and saved in the software.  Or " & _
            "alternatively, you can resupply the discharge record."
        ProtectThisSheet 1
        Application.ScreenUpdating = True
    End If
    
End Sub

Sub obCMSobCFSClickedOnMyInput()
    Dim cc As String, i As Long, dmy As Double
    If EnableOnError Then On Error Resume Next
    cc = Right(Cells(1, 1).Value, 5)
    
    UnprotectThisSheet 1
    Range("A2:A" & Rows.Count).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ProtectThisSheet 1
    Cells(2, 1).Select
    
    If Cells(1, 10).Value = 1 Then 'cms
        If cc = "(cfs)" Then
            Cells(1, 1).Value = "Discharge (cms)"
            i = 1
            Do While Not IsEmpty(Cells(i + 1, 1))
                i = i + 1
                If IsNumeric(Cells(i, 1).Value) Then
                    dmy = Cells(i, 1).Value
                    Cells(i, 1).Value = Format(dmy * 0.3048 ^ 3, "###0.##")
                End If
            Loop
        End If
    End If
    If Cells(1, 10).Value = 2 Then 'cfs
        If cc = "(cms)" Then
            Cells(1, 1).Value = "Discharge (cfs)"
            i = 1
            Do While Not IsEmpty(Cells(i + 1, 1))
                i = i + 1
                If IsNumeric(Cells(i, 1).Value) Then
                    dmy = Cells(i, 1).Value
                    Cells(i, 1).Value = Format(dmy / 0.3048 ^ 3, "###0.##")
                End If
            Loop
        End If
    End If
End Sub

Private Sub ManipulateCrossSection()
    Dim InSt As Worksheet, xRange As Range, yRange As Range
    Dim xx(1001) As Double, yY(1001) As Double
    Dim Node As Long, i As Long, j As Long
    Dim Xlf As Double, Xrf As Double
    Dim Ylf As Double, Yrf As Double
    Dim Depth As Double, Ymin As Double, X As Double, Y As Double, yWS As Double
    Dim Pl As Double, Al As Double, Rl As Double ' left floodplain
    Dim Pr As Double, Ar As Double, Rr As Double ' right floodplain
    Dim Pm As Double, Am As Double, Rm As Double ' main channel
    
    If EnableOnError Then On Error Resume Next
    Set InSt = Worksheets("Input")
    
    If InSt.Cells(1, 1).Value = "XS" Then
        MessageOnWelcome "Calculating information about channel geometry." & _
            vbLf & vbLf & "Please wait..."
        Node = 0
        Do While Not IsEmpty(InSt.Cells(Node + 1, 3))
            Node = Node + 1
        Loop
        
        If InSt.Cells(1, 3).Value > InSt.Cells(Node, 3).Value Then 'reordering XS
            For i = 1 To Node / 2
                Xlf = InSt.Cells(i, 3).Value
                Ylf = InSt.Cells(i, 4).Value
                InSt.Cells(i, 3).Value = InSt.Cells(Node + 1 - i, 3).Value
                InSt.Cells(i, 4).Value = InSt.Cells(Node + 1 - i, 4).Value
                InSt.Cells(Node + 1 - i, 3).Value = Xlf
                InSt.Cells(Node + 1 - i, 4).Value = Ylf
            Next
        End If
        
        Set xRange = Range(InSt.Cells(1, 3), InSt.Cells(Node, 3))
        Set yRange = Range(InSt.Cells(1, 4), InSt.Cells(Node, 4))
        
        If InSt.Cells(2, 1).Value = "Yes" Then
            Xlf = InSt.Cells(4, 1).Value
            If Xlf < Application.WorksheetFunction.Min(xRange) Then _
                Xlf = Application.WorksheetFunction.Min(xRange)
        Else
            Xlf = Application.WorksheetFunction.Min(xRange)
        End If
        If InSt.Cells(2, 2).Value = "Yes" Then
            Xrf = InSt.Cells(4, 2).Value
            If Xrf > Application.WorksheetFunction.Max(xRange) Then _
                Xrf = Application.WorksheetFunction.Max(xRange)
        Else
            Xrf = Application.WorksheetFunction.Max(xRange)
        End If
        
        InSt.Cells(52, 9).Value = Xlf - Application.WorksheetFunction.Min(xRange)
            ' = left floodplain width
        InSt.Cells(52, 10).Value = Xrf - Xlf ' = main channel width
        InSt.Cells(52, 11).Value = Application.WorksheetFunction.Max(xRange) - Xrf
            ' = right floodplain width
        
        For i = 1 To xRange.Cells.Count - 1
            If xRange.Cells(i) <= Xlf And xRange.Cells(i + 1) >= Xlf Then
                Ylf = yRange.Cells(i) + (yRange.Cells(i + 1) - yRange.Cells(i)) / _
                    (xRange.Cells(i + 1) - xRange.Cells(i)) * _
                    (Xlf - xRange.Cells(i))
            End If
            Exit For
        Next
        For i = 1 To xRange.Cells.Count - 1
            If (xRange.Cells(i) <= Xrf And xRange.Cells(i + 1) >= Xrf) Then
                Yrf = yRange.Cells(i) + (yRange.Cells(i + 1) - yRange.Cells(i)) / _
                    (xRange.Cells(i + 1) - xRange.Cells(i)) * _
                    (Xrf - xRange.Cells(i))
            End If
            Exit For
        Next
        
        Ymin = Application.WorksheetFunction.Min(yRange) ' min. elevation
        Depth = Application.WorksheetFunction.Max(yRange) - Ymin ' max depth
        
        xx(0) = xRange.Cells(1)
        xx(1000) = xRange.Cells(Node)
        yY(0) = yRange.Cells(1)
        yY(1000) = yRange.Cells(Node)
        For i = 1 To 999
            xx(i) = xx(0) + i * (xx(1000) - xx(0)) / 1000
        Next
        j = 1
        For i = 1 To 999
            Do While xRange(j).Value < xx(i)
                j = j + 1
            Loop
            yY(i) = yRange.Cells(j - 1) + (yRange.Cells(j) - yRange.Cells(j - 1)) / _
                (xRange.Cells(j) - xRange.Cells(j - 1)) * _
                (xx(i) - xRange.Cells(j - 1))
        Next

        For j = 10 To 15
            InSt.Cells(1, j).Value = 0
        Next
        For i = 2 To 51 ' define depth
            InSt.Cells(i, 9).Value = (i - 1) * Depth / 50
            yWS = InSt.Cells(i, 9).Value + Ymin
            Al = 0
            Am = 0
            Ar = 0
            Pl = 0
            Pm = 0
            Pr = 0
            For j = 0 To 999
                X = 0.5 * (xx(j) + xx(j + 1))
                Y = 0.5 * (yY(j) + yY(j + 1))
                If yWS > Y Then ' segment is under water
                    If X < Xlf Then ' on left floodplain
                        Al = Al + (yWS - Y) * (xx(j + 1) - xx(j))
                        Pl = Pl + ((yY(j + 1) - yY(j)) ^ 2 + (xx(j + 1) - xx(j)) ^ 2) ^ 0.5
                    ElseIf X <= Xrf Then 'on main channel
                        Am = Am + (yWS - Y) * (xx(j + 1) - xx(j))
                        Pm = Pm + ((yY(j + 1) - yY(j)) ^ 2 + (xx(j + 1) - xx(j)) ^ 2) ^ 0.5
                    Else ' on right floodplain
                        Ar = Ar + (yWS - Y) * (xx(j + 1) - xx(j))
                        Pr = Pr + ((yY(j + 1) - yY(j)) ^ 2 + (xx(j + 1) - xx(j)) ^ 2) ^ 0.5
                    End If
                End If
            Next
            If Pl > 0 Then
                Rl = Al / Pl
            Else
                Rl = 0
            End If
            If Pm > 0 Then
                Rm = Am / Pm
            Else
                Rm = 0
            End If
            If Pr > 0 Then
                Rr = Ar / Pr
            Else
                Rr = 0
            End If
            If Rm > InSt.Cells(i - 1, 10).Value Then
                InSt.Cells(i, 10).Value = Rm
            Else
                InSt.Cells(i, 10).Value = InSt.Cells(i - 1, 10).Value
            End If
            InSt.Cells(i, 11).Value = Am
            If Rl > InSt.Cells(i - 1, 12).Value Then
                InSt.Cells(i, 12).Value = Rl
            Else
                InSt.Cells(i, 12).Value = InSt.Cells(i - 1, 12).Value
            End If
            InSt.Cells(i, 13).Value = Al
            If Rr > InSt.Cells(i - 1, 14).Value Then
                InSt.Cells(i, 14).Value = Rr
            Else
                InSt.Cells(i, 14).Value = InSt.Cells(i - 1, 14).Value
            End If
            InSt.Cells(i, 15).Value = Ar
        Next
        EndMessageOnWelcome 1
    End If
    i = MyMsgBox("Input data are complete!  You can save the input data into a " & _
        "project file by clicking ""File - Save BAGS project"" on the menu " & _
        "bar.  Saving the project will allow you to reload the data at a later date " & _
        "(by clicking ""File - Open BAGS project"" on the menu).  We strongly " & _
        "suggest that you save the project before proceeding to the next task." _
        & vbLf & vbLf & _
        "Save project now?", "Save project", vbYes)
    If i = vbCancel Then Exit Sub
    If i = vbYes Then
        UserFormInUse = True
        BAGSModule2.SaveMyProject 1
        Do While UserFormInUse
            DoEvents
        Loop
    End If
    SurfaceBasedParker90 1 ' proceed to bedload calculation on BAGSModule3
End Sub

Sub InitializingProgressBar(dmy As Integer)
    UnprotectThisSheet 1
    ActiveSheet.Shapes("ProgressBarBackground").Visible = True
    ActiveSheet.Shapes("ProgressBarForeground").Left = _
        ActiveSheet.Shapes("ProgressBarBackground").Left
    ActiveSheet.Shapes("ProgressBarForeground").Top = _
        ActiveSheet.Shapes("ProgressBarBackground").Top
    ActiveSheet.Shapes("ProgressBarForeground").Height = _
        ActiveSheet.Shapes("ProgressBarBackground").Height
    ActiveSheet.Shapes("ProgressBarForeground").Width = 0
    ActiveSheet.Shapes("ProgressBarForeground").Visible = True
    ProtectThisSheet 1
End Sub

Sub UpdatingProgressBar(dmy As Integer)
    UnprotectThisSheet 1
    ActiveSheet.Shapes("ProgressBarForeground").Width = _
        ActiveSheet.Shapes("ProgressBarForeground").Width + _
        ActiveSheet.Shapes("ProgressBarBackground").Width / 100
    If ActiveSheet.Shapes("ProgressBarForeground").Width > ActiveSheet.Shapes("ProgressBarBackground").Width Then _
        ActiveSheet.Shapes("ProgressBarForeground").Width = 0
    ProtectThisSheet 1
End Sub

Sub PlotFloodplains(dmy As Long)
    Dim LeftF As Double, RightF As Double
    Dim InSt As Worksheet

    Set InSt = Worksheets("Input")
    
    If IsEmpty(Cells(23, 7)) Then
        LeftF = -1E+20
    Else
        LeftF = Cells(23, 7).Value
    End If
    If IsEmpty(Cells(24, 7)) Then
        RightF = 1E+20
    Else
        RightF = Cells(24, 7).Value
    End If
    
    dmy = 0
    Do While Not IsEmpty(InSt.Cells(dmy + 1, 3))
        dmy = dmy + 1
        If InSt.Cells(dmy, 3).Value <= LeftF Or _
            InSt.Cells(dmy, 3).Value >= RightF Then
            Cells(dmy, 10).Value = InSt.Cells(dmy, 4).Value
            Cells(dmy, 11).Value = InSt.Cells(dmy, 3).Value
        ElseIf dmy > 1 Then
            If InSt.Cells(dmy - 1, 3).Value <= LeftF Then
                Cells(dmy, 11).Value = LeftF
                Cells(dmy, 10).Value = InSt.Cells(dmy - 1, 4) + _
                    (InSt.Cells(dmy, 4) - InSt.Cells(dmy - 1, 4)) / _
                    (InSt.Cells(dmy, 3) - InSt.Cells(dmy - 1, 3)) * _
                    (LeftF - InSt.Cells(dmy - 1, 3))
            ElseIf InSt.Cells(dmy + 1, 3).Value >= RightF Then
                Cells(dmy, 11).Value = RightF
                Cells(dmy, 10).Value = InSt.Cells(dmy + 1, 4) + _
                    (InSt.Cells(dmy, 4) - InSt.Cells(dmy + 1, 4)) / _
                    (InSt.Cells(dmy, 3) - InSt.Cells(dmy + 1, 3)) * _
                    (RightF - InSt.Cells(dmy + 1, 3))
            Else
                Cells(dmy, 10).ClearContents
                Cells(dmy, 11).ClearContents
            End If
        Else
            Cells(dmy, 10).ClearContents
            Cells(dmy, 11).ClearContents
        End If
    Loop
    
End Sub

Sub AgreementClicked()
    Worksheets("Agreement").Visible = True
    Worksheets("Agreement").Select
    Worksheets("Welcome").Visible = False
End Sub

Sub AgreeOnAgreementClicked()
    Worksheets("Welcome").Visible = True
    Worksheets("Welcome").Select
    Worksheets("Agreement").Visible = False
    RunSoftware 1
End Sub

Sub DisagreeOnAgreementClicked()
    ThisWorkbook.Close
End Sub


