Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

Function ParkerKlingemanBasicEquation(Beta As Double, Taur50 As Double, R As Double, _
    g As Double, Dk As Double, D50mm As Double, Width As Double, Qw As Double, f() As Double, _
    p() As Double, Psi() As Double, Qs As Double, H As Double, Phi50 As Double, _
    Slope As Double, Nsize As Long, EqOption As String) As Integer 'Nsize is number of fractions
        'EqOption = "PK" or "Bakke"
        
    Dim ExitFunction As Label
    Dim i As Long
    Dim Ustar As Double, Rough As Double, D50 As Double
    Dim Di As Double, dmy As Double
    Dim nn As Double, nD As Double
    Dim tau As Double, tauT As Double
    Dim Rho As Double
    
    Rho = 1000 'doesn't matter but given a correct value anyway
    
    ParkerKlingemanBasicEquation = 1
    
    D50 = D50mm / 1000
    Rough = Dk * D50
    
' Revised in 2006 to include roughness correction
    If EqOption = "PK" And Worksheets("Input").Cells(19, 1).Value Then
        nn = Worksheets("Input").Cells(19, 2).Value
        nD = 0.04 * Rough ^ (1 / 6)
        If nD <= nn Then 'apply correction
            H = (nn * Qw / Width / Slope ^ 0.5) ^ (3 / 5)
            tauT = Rho * g * H * Slope
            tau = tauT * (nD / nn) ^ 1.5
            Ustar = (tau / Rho) ^ 0.5
            Worksheets("Input").Cells(19, 2).Interior.ColorIndex = xlNone
        Else 'original method
            Worksheets("Input").Cells(19, 2).Interior.ColorIndex = 36
            H = Width / 100 ' initial guess
            If CalculateShearVelocityWithConstantWidth(g, Qw, Width, H, Slope, Rough, Ustar) = 0 Then
                ParkerKlingemanBasicEquation = 0
                Exit Function
            End If
        End If
    Else 'original method
        H = Width / 100 ' initial guess
        If CalculateShearVelocityWithConstantWidth(g, Qw, Width, H, Slope, Rough, Ustar) = 0 Then
            ParkerKlingemanBasicEquation = 0
            Exit Function
        End If
    End If
' 2006 revision ends here

    Qs = 0
    For i = 1 To Nsize
        Di = 2 ^ (0.5 * (Psi(i) + Psi(i + 1))) / 1000
        Phi50 = Ustar ^ 2 / R / g / D50 / Taur50
'        dmy = 1 - 0.853 / Phi50 * (Di / D50) ^ Beta
'        If dmy > 1 Then
'            p(i) = 11.2 * (dmy) ^ 4.5 * f(i)
'        Else
'            p(i) = 0
'        End If
        'the following code is added to replace the code above for convergence purposes
        dmy = Phi50 / (Di / D50) ^ Beta
        If PKorPKM = "PK" Then
            If dmy > 0.95 Then
                p(i) = 11.2 * (1 - 0.853 / dmy) ^ 4.5 * f(i)
            Else
                p(i) = 0.00242947 * dmy ^ 35.71387887 * f(i)
            End If
        ElseIf PKorPKM = "PKM" Then
            p(i) = GinParker82(dmy) * f(i)
        Else
            MsgBox "No equation is selected!"
            Exit Function
        End If
        Qs = Qs + p(i)
    Next
    If Qs > 0 Then
        For i = 1 To Nsize
            p(i) = p(i) / Qs
        Next
        Qs = Qs * Width * Ustar ^ 3 / R / g
    End If
    
    Exit Function
ExitFunction:
    ParkerKlingemanBasicEquation = 0
End Function


Function ParkerklingemanEquationWithCrossSection(Beta As Double, Taur50 As Double, _
    R As Double, g As Double, Dk As Double, D50mm As Double, Qw As Double, f() As Double, _
    p() As Double, Psi() As Double, Qs As Double, H As Double, Phi50 As Double, _
    Slope As Double, Nsize As Long, EqOption As String) As Integer 'Nsize is number of fractions
        'EqOption = "PK" or "Bakke"
        
    Dim i As Long
    Dim Ustar As Double, Rh As Double, Area As Double, Rough As Double
    Dim Mn1 As Double, Mn2 As Double ' Manning's n for floodplains
    Dim Rh1 As Double, Rh2 As Double ' hydraulic radius for floodplains
    Dim Area1 As Double, Area2 As Double ' flow area in floodplain
    Dim Qw1 As Double, Qw2 As Double, Qwc As Double ' discharge in floodplains and the main channel
    Dim Di As Double, D50 As Double, dmy As Double
    Dim CheckOut As Label
    Dim nn As Double, nD As Double
    Dim tau As Double, tauT As Double
    Dim Rho As Double
    
    Rho = 1000 'doesn't matter but given a correct value anyway
    
    If OnErrorOn Then On Error GoTo CheckOut
    
    D50 = D50mm / 1000
    Rough = Dk * D50
    Mn1 = Worksheets("Input").Cells(3, 1).Value
    Mn2 = Worksheets("Input").Cells(3, 2).Value
    
' Revised in 2006 to include roughness correction
    If EqOption = "PK" And Worksheets("Input").Cells(19, 1).Value Then
        nn = Worksheets("Input").Cells(19, 2).Value
        nD = 0.04 * Rough ^ (1 / 6)
        If nD <= nn Then 'apply correction
            GetDepthWithDischargeManningsn Qw, Slope, nn, Mn1, Mn2, Qwc, Qw1, Qw2, H, Rh, Area
            tauT = Rho * g * Rh * Slope
            tau = tauT * (nD / nn) ^ 1.5
            Ustar = (tau / Rho) ^ 0.5
            Worksheets("Input").Cells(19, 2).Interior.ColorIndex = xlNone
        Else 'original method
            Worksheets("Input").Cells(19, 2).Interior.ColorIndex = 36
            GetDepthWithDischarge Qw, Slope, Rough, Mn1, Mn2, Qwc, Qw1, Qw2, H, Rh, Area, g
            Ustar = (g * Rh * Slope) ^ 0.5
        End If
    Else 'original method
        GetDepthWithDischarge Qw, Slope, Rough, Mn1, Mn2, Qwc, Qw1, Qw2, H, Rh, Area, g
        Ustar = (g * Rh * Slope) ^ 0.5
    End If
' 2006 revision ends here
       
    Phi50 = Ustar ^ 2 / R / g / D50 / Taur50
    
    Qs = 0
    For i = 1 To Nsize
        Di = 2 ^ (0.5 * (Psi(i) + Psi(i + 1))) / 1000
'        dmy = 1 - 0.853 / Phi50 * (Di / D50) ^ Beta
'        If dmy > 0 Then
'            p(i) = 11.2 * (dmy) ^ 4.5 * f(i)
'        Else
'            p(i) = 0
'        End If
        'the following code is added to replace the code above for convergence purposes
        dmy = Phi50 / (Di / D50) ^ Beta
        If PKorPKM = "PK" Then
            If dmy > 0.95 Then
                p(i) = 11.2 * (1 - 0.853 / dmy) ^ 4.5 * f(i)
            Else
                p(i) = 0.00242947 * dmy ^ 35.71387887 * f(i)
            End If
        ElseIf PKorPKM = "PKM" Then
            p(i) = GinParker82(dmy) * f(i)
        Else
            MsgBox "No equation is selected!"
            Exit Function
        End If
        Qs = Qs + p(i)
    Next
    If Qs > 0 Then
        For i = 1 To Nsize
            p(i) = p(i) / Qs
        Next
        Qs = Qs * Ustar * Slope * Area / R
    End If
    ParkerklingemanEquationWithCrossSection = 1
    Exit Function
    
CheckOut:
    MsgBox "An error occured while ""ParkerklingemanEquationWithCrossSection"" is executed!"
    ParkerklingemanEquationWithCrossSection = 0
End Function

Sub PresentResultsForParkerKlingeman82(Qs() As Double, Phi50() As Double, H() As Double, p() As Double, _
    MyOption As Integer)
    
    'MyOption = 0 for ParkerKlingeman82 and MyOption = 1 for Bakke et al. (1999)
    
    Dim i As Long, j As Long, Nsize As Long, dmy As Double
    Dim nXS As Long
    Dim MyBk As Workbook, InSt As Worksheet, StSt As Worksheet
    Dim Rh As Double, Area As Double, Rh1 As Double, Area1 As Double, Rh2 As Double, Area2 As Double
    Dim MySize As Range, MyFiner As Range, ChD(10) As Double
    Dim xRange As Range, yRange As Range
    Dim SampleStarts As Long, SampleEnds As Long
    Dim cc As String
    
    Set InSt = Worksheets("Input"): Set StSt = Worksheets("Storage")
    
    Nsize = 0
    Do While Not IsEmpty(InSt.Cells(Nsize + 2, 7))
        Nsize = Nsize + 1
    Loop
    
    'add workbook with two worksheets
    i = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = 3
    Set MyBk = Workbooks.Add
    Application.SheetsInNewWorkbook = i

    ThisWorkbook.Activate
    
    MyBk.Sheets(1).Name = "Note"
    MyBk.Sheets(2).Name = "Input"
    MyBk.Sheets(3).Name = "Output"
    
    MyBk.Sheets(1).Cells.Interior.ColorIndex = 2
    MyBk.Sheets(1).Cells(2, 2).Value = _
        "This workbook contains bedload transport calculation results from USDA Forest Service's BAGS software."
    If MyOption = 0 Then _
        MyBk.Sheets(1).Cells(4, 2).Value = _
            "Bedload transport equation used: The substrate-based bedload equation of Parker and Klingeman (1982)."
    If MyOption = 1 And PKorPKM = "PK" Then
        MyBk.Sheets(1).Cells(4, 2).Value = _
            "Bedload transport equation used: The substrate-based bedload equation of Parker and Klingeman (1982)"
        MyBk.Sheets(1).Cells(5, 2).Value = _
            "      with the calibration procedure of Bakke et al. (1999)."
        MyBk.Sheets(1).Cells(7, 2).Value = "Optimization Parameters:"
        MyBk.Sheets(1).Cells(8, 3).Value = "Taur50 = " & InSt.Cells(15, 1).Value
        MyBk.Sheets(1).Cells(9, 3).Value = "Exponent = " & InSt.Cells(15, 2).Value
    End If
    If MyOption = 1 And PKorPKM = "PKM" Then
        MyBk.Sheets(1).Cells(4, 2).Value = _
            "Bedload transport equation used: The substrate-based bedload equation of Parker, Klingeman, and"
        MyBk.Sheets(1).Cells(5, 2).Value = _
            "      McLean (982) with the calibration procedure of Bakke et al. (1999)."
        MyBk.Sheets(1).Cells(7, 2).Value = "Optimization Parameters:"
        MyBk.Sheets(1).Cells(8, 3).Value = "Taur50 = " & InSt.Cells(16, 1).Value
        MyBk.Sheets(1).Cells(9, 3).Value = "Exponent = " & InSt.Cells(16, 2).Value
    End If
    MyBk.Sheets(1).Cells(11, 2).Value = _
        "Input data are stored in worksheet ""Input"" and results are stored in worksheet ""Output""."
    MyBk.Sheets(1).Cells(12, 2).Value = _
        "Calculation was performed by " & Application.UserName & " on " & Date & "."

    MyBk.Sheets(2).Columns("A:A").ColumnWidth = 2
    MyBk.Sheets(2).Columns("B:B").ColumnWidth = 24
    MyBk.Sheets(2).Columns("C:C").ColumnWidth = 12
    MyBk.Sheets(2).Columns("D:D").ColumnWidth = 2
    MyBk.Sheets(2).Columns("E:E").ColumnWidth = 15
    MyBk.Sheets(2).Columns("F:F").ColumnWidth = 15
    MyBk.Sheets(2).Columns("G:G").ColumnWidth = 2
    MyBk.Sheets(2).Columns("H:H").ColumnWidth = 12
    MyBk.Sheets(2).Columns("I:I").ColumnWidth = 12
    MyBk.Sheets(2).Cells.Interior.ColorIndex = 2
    MyBk.Sheets(2).Cells.HorizontalAlignment = xlCenter
    MyBk.Sheets(2).Cells(2, 5).HorizontalAlignment = xlGeneral
    MyBk.Sheets(2).Cells(2, 8).HorizontalAlignment = xlGeneral
    
    ' input date: slope
    If InSt.Cells(5, 1).Value = "W.S." Then _
        MyBk.Sheets(2).Cells(2, 2).Value = "Water surface slope"
    If InSt.Cells(5, 1).Value = "Bed" Then _
        MyBk.Sheets(2).Cells(2, 2).Value = "Channel bed slope"
    If InSt.Cells(5, 1).Value = "Model" Then _
        MyBk.Sheets(2).Cells(2, 2).Value = "Friction slope from a model"
    MyBk.Sheets(2).Cells(2, 3).Value = InSt.Cells(5, 2).Value

    ' input data: Bankfull width or cross section
    MyBk.Sheets(2).Cells(4, 2).Value = "Bankfull width"
    If InSt.Cells(1, 1).Value = "XS" Then 'cross section
        MyBk.Sheets(2).Cells(4, 3).Value = "N/A"
        
        MyBk.Sheets(2).Cells(9, 2).Value = "Left floodplain boundary"
        MyBk.Sheets(2).Cells(10, 2).Value = "Left floodplain Manning's n"
        If InSt.Cells(2, 1).Value = "Yes" Then
            MyBk.Sheets(2).Cells(9, 3).Value = InSt.Cells(4, 1).Value
            MyBk.Sheets(2).Cells(9, 3).NumberFormat = "###0.##"" m"""
            MyBk.Sheets(2).Cells(10, 3).Value = InSt.Cells(3, 1).Value
        Else
            MyBk.Sheets(2).Cells(9, 3).Value = "N/A"
            MyBk.Sheets(2).Cells(10, 3).Value = "N/A"
        End If
        MyBk.Sheets(2).Cells(11, 2).Value = "Right floodplain boundary"
        MyBk.Sheets(2).Cells(12, 2).Value = "Right floodplain Manning's n"
        If InSt.Cells(2, 2).Value = "Yes" Then
            MyBk.Sheets(2).Cells(11, 3).Value = InSt.Cells(4, 2).Value
            MyBk.Sheets(2).Cells(11, 3).NumberFormat = "###0.##"" m"""
            MyBk.Sheets(2).Cells(12, 3).Value = InSt.Cells(3, 2).Value
        Else
            MyBk.Sheets(2).Cells(11, 3).Value = "N/A"
            MyBk.Sheets(2).Cells(12, 3).Value = "N/A"
        End If
        MyBk.Sheets(2).Cells(14, 2).Value = "CROSS SECTION"
        MyBk.Sheets(2).Cells(15, 2).Value = "Lateral distance (m)"
        MyBk.Sheets(2).Cells(15, 3).Value = "Elevation (m)"
        i = 0
        Do While Not IsEmpty(InSt.Cells(i + 1, 3))
            i = i + 1
            MyBk.Sheets(2).Cells(i + 15, 2).Value = InSt.Cells(i, 3).Value
            MyBk.Sheets(2).Cells(i + 15, 3).Value = InSt.Cells(i, 4).Value
        Loop
        nXS = i
    Else 'bankfull width
        MyBk.Sheets(2).Cells(4, 3).Value = InSt.Cells(1, 2).Value
        MyBk.Sheets(2).Cells(4, 3).NumberFormat = "###0.##"" m"""
    End If

    ' input data: Discharge or flow duration curve
    If InSt.Cells(6, 1).Value = "(A)" Then 'single discharge
        MyBk.Sheets(2).Cells(6, 2).Value = "Water discharge"
        MyBk.Sheets(2).Cells(6, 3).Value = InSt.Cells(1, 16).Value
        MyBk.Sheets(2).Cells(6, 3).NumberFormat = "###0.##"" cms"""
    ElseIf InSt.Cells(6, 1).Value = "(B)" Then 'minimum and maximum discharge
        MyBk.Sheets(2).Cells(6, 2).Value = "Min. water discharge"
        MyBk.Sheets(2).Cells(6, 3).Value = InSt.Cells(1, 16).Value
        MyBk.Sheets(2).Cells(6, 3).NumberFormat = "###0.##"" cms"""
        MyBk.Sheets(2).Cells(7, 2).Value = "Max. water discharge"
        MyBk.Sheets(2).Cells(7, 3).Value = InSt.Cells(26, 16).Value
        MyBk.Sheets(2).Cells(7, 3).NumberFormat = "###0.##"" cms"""
    Else 'flow duration curve
        MyBk.Sheets(2).Cells(6, 2).Value = "Flow duration curve is given"
        MyBk.Sheets(2).Cells(7, 2).Value = "on Columns E and F"
        MyBk.Sheets(2).Cells(2, 5).Value = "FLOW DURATION CURVE"
        MyBk.Sheets(2).Cells(4, 5).Value = "Discharge (cms)"
        MyBk.Sheets(2).Cells(3, 6).Value = "Exceedance"
        MyBk.Sheets(2).Cells(4, 6).Value = "probability (%)"
        For i = 1 To 26
            MyBk.Sheets(2).Cells(i + 4, 5).Value = InSt.Cells(i, 16).Value
            MyBk.Sheets(2).Cells(i + 4, 5).NumberFormat = "###0.##"
            MyBk.Sheets(2).Cells(i + 4, 6).Value = InSt.Cells(i, 17).Value
            MyBk.Sheets(2).Cells(i + 4, 6).NumberFormat = "###0.##"
        Next
    End If
    
    ' input date: substrate grain size distribution
    MyBk.Sheets(2).Cells(2, 8).Value = "SUBSTRATE GRAIN SIZE DISTRIBUTION"
    MyBk.Sheets(2).Cells(3, 8).Value = "Size (mm)"
    MyBk.Sheets(2).Cells(3, 9).Value = "% Finer"
    For i = 1 To Nsize + 1
        MyBk.Sheets(2).Cells(i + 3, 8).Value = Format(InSt.Cells(i, 7).Value, "###0.##")
        MyBk.Sheets(2).Cells(i + 3, 9).Value = Format(InSt.Cells(i, 8).Value, "###0.##")
    Next
    
    Set MySize = Range(InSt.Cells(1, 7), InSt.Cells(Nsize + 1, 7))
    Set MyFiner = Range(InSt.Cells(1, 8), InSt.Cells(Nsize + 1, 8))
    GetGrainSizeStatistics Nsize, MySize, MyFiner, ChD
    
    MyBk.Sheets(2).Cells(Nsize + 7, 8).Value = "STATISTICS OF THE ABOVE GRAIN SIZE DISTRIBUTION:"
    MyBk.Sheets(2).Cells(Nsize + 8, 8).Value = "Geometric mean (mm)"
    MyBk.Sheets(2).Cells(Nsize + 9, 8).Value = "Geometric standard deviation"
    MyBk.Sheets(2).Cells(Nsize + 10, 8).Value = "D10 (mm)"
    MyBk.Sheets(2).Cells(Nsize + 11, 8).Value = "D16 (mm)"
    MyBk.Sheets(2).Cells(Nsize + 12, 8).Value = "D25 (mm)"
    MyBk.Sheets(2).Cells(Nsize + 13, 8).Value = "D50 (mm)"
    MyBk.Sheets(2).Cells(Nsize + 14, 8).Value = "D65 (mm)"
    MyBk.Sheets(2).Cells(Nsize + 15, 8).Value = "D75 (mm)"
    MyBk.Sheets(2).Cells(Nsize + 16, 8).Value = "D84 (mm)"
    MyBk.Sheets(2).Cells(Nsize + 17, 8).Value = "D90 (mm)"
    Range(MyBk.Sheets(2).Cells(Nsize + 7, 8), MyBk.Sheets(2).Cells(Nsize + 17, 8)).HorizontalAlignment = xlGeneral
    For i = 0 To 9
        MyBk.Sheets(2).Cells(Nsize + 8 + i, 10).Value = Format(ChD(i), "###0.##")
    Next
    ' input data
    If MyOption = 1 Then 'bedload sampling data
        MyBk.Sheets(2).Cells(Nsize + 20, 8).Value = "BEDLOAD SAMPLING DATA"
        MyBk.Sheets(2).Cells(Nsize + 20, 8).HorizontalAlignment = xlGeneral
        MyBk.Sheets(2).Cells(Nsize + 21, 8).Value = "Discharge"
        MyBk.Sheets(2).Cells(Nsize + 22, 8).Value = "(cms)"
        MyBk.Sheets(2).Cells(Nsize + 21, 9).Value = "Bedload"
        MyBk.Sheets(2).Cells(Nsize + 22, 9).Value = "(kg/min.)"
        MyBk.Sheets(2).Cells(Nsize + 21, 10).Value = "D50"
        MyBk.Sheets(2).Cells(Nsize + 22, 10).Value = "(mm)"
        i = 0
        Do While Not IsEmpty(InSt.Cells(i + 1, 18))
            i = i + 1
            MyBk.Sheets(2).Cells(Nsize + 22 + i, 8).Value = InSt.Cells(i, 18).Value
            MyBk.Sheets(2).Cells(Nsize + 22 + i, 9).Value = InSt.Cells(i, 19).Value
            MyBk.Sheets(2).Cells(Nsize + 22 + i, 10).Value = StSt.Cells(i, 5).Value
            If InSt.Cells(i, 40).Value Then _
                MyBk.Sheets(2).Cells(Nsize + 22 + i, 11).Value = "Outlier"
        Loop
        SampleStarts = Nsize + 23
        SampleEnds = Nsize + 22 + i
    Else 'main channel Manning's n
        If Worksheets("Input").Cells(19, 1).Value Then
            MyBk.Sheets(2).Cells(Nsize + 19, 8).Value = "Main channel Manning's n"
            MyBk.Sheets(2).Cells(Nsize + 19, 8).HorizontalAlignment = xlGeneral
            MyBk.Sheets(2).Cells(Nsize + 19, 10).Value = _
                Worksheets("Input").Cells(19, 2).Value
            If Worksheets("Input").Cells(19, 2).Interior.ColorIndex <> xlNone Then
                    MyBk.Sheets(2).Cells(Nsize + 20, 8).Value = "(This main channel Manning's n is not used because it is"
                    MyBk.Sheets(2).Cells(Nsize + 20, 8).HorizontalAlignment = xlGeneral
                    MyBk.Sheets(2).Cells(Nsize + 21, 8).Value = "smaller than what is calculated based on grain roughness.)"
                    MyBk.Sheets(2).Cells(Nsize + 21, 8).HorizontalAlignment = xlGeneral
            End If
        End If
    End If
    'end for input data
    
    MyBk.Sheets(3).Columns("A:A").ColumnWidth = 2
    MyBk.Sheets(3).Columns("B:B").ColumnWidth = 10
    MyBk.Sheets(3).Columns("C:C").ColumnWidth = 10
    MyBk.Sheets(3).Columns("D:D").ColumnWidth = 10
    MyBk.Sheets(3).Columns("G:G").ColumnWidth = 2
    MyBk.Sheets(3).Columns("H:H").ColumnWidth = 10
    MyBk.Sheets(3).Columns("I:I").ColumnWidth = 15
    MyBk.Sheets(3).Columns("J:J").ColumnWidth = 15
    MyBk.Sheets(3).Columns("K:K").ColumnWidth = 10
    MyBk.Sheets(3).Cells.Interior.ColorIndex = 2
    MyBk.Sheets(3).Cells.HorizontalAlignment = xlCenter
    MyBk.Sheets(3).Cells(2, 2).HorizontalAlignment = xlGeneral
    MyBk.Sheets(3).Cells(5, 2).HorizontalAlignment = xlGeneral
    MyBk.Sheets(3).Cells(4, 8).HorizontalAlignment = xlGeneral
    MyBk.Sheets(3).Cells(5, 14).HorizontalAlignment = xlGeneral
    
    ' bedload transport rate
    If InSt.Cells(6, 1).Value = "(A)" Then 'single discharge
        MyBk.Sheets(3).Cells(2, 2).Value = "Bedload transport rate (kg/min.)"
        If Qs(1) > 0 Then _
            MyBk.Sheets(3).Cells(2, 6).Value = Qs(1) * 2650 * 60 ' m3/s to kg/min.
        MyBk.Sheets(3).Cells(3, 2).Value = "Normalized Shields stress"
        MyBk.Sheets(3).Cells(3, 2).HorizontalAlignment = xlGeneral
        MyBk.Sheets(3).Cells(3, 6).Value = Phi50(1) ' m3/s to kg/min.
    Else
        If InSt.Cells(6, 1).Value = "(B)" Then 'min. and max. discharge
            MyBk.Sheets(3).Cells(2, 2).Value = "Rating curves are presented starting Column H"
        Else 'duration curve
            MyBk.Sheets(3).Cells(2, 2).Value = "Average bedload transport rate (kg/min.)"
            Qs(0) = 0
            For i = 1 To 25
                Qs(0) = Qs(0) + 0.5 * (Qs(i) + Qs(i + 1)) * _
                    Abs(InSt.Cells(i + 1, 17).Value - InSt.Cells(i, 17).Value) / 100
            Next
            MyBk.Sheets(3).Cells(2, 6).Value = Qs(0) * 2650 * 60 ' m3/s to kg/min.
        End If
        MyBk.Sheets(3).Cells(4, 8).Value = "RATING CURVES"
        MyBk.Sheets(3).Cells(5, 8).Value = "Discharge"
        MyBk.Sheets(3).Cells(6, 8).Value = "(cms)"
        MyBk.Sheets(3).Cells(5, 9).Value = "Bedload transport"
        MyBk.Sheets(3).Cells(6, 9).Value = "rate (kg/min.)"
        MyBk.Sheets(3).Cells(5, 10).Value = "Transport" '"Normalized"
        MyBk.Sheets(3).Cells(6, 10).Value = "Stage" '"Shields stress"
        MyBk.Sheets(3).Cells(5, 11).Value = "D50"
        MyBk.Sheets(3).Cells(6, 11).Value = "(mm)"
        If InSt.Cells(1, 1).Value = "XS" Then
            MyBk.Sheets(3).Cells(5, 12).Value = "Max water"
            MyBk.Sheets(3).Cells(5, 13).Value = "Hydraulic"
            MyBk.Sheets(3).Cells(6, 13).Value = "radius (m)"
        Else
            MyBk.Sheets(3).Cells(5, 12).Value = "Water"
        End If
        MyBk.Sheets(3).Cells(6, 12).Value = "depth (m)"
        MyBk.Sheets(3).Cells(5, 14).Value = "Sediment transport rate by size, in kg/min."
        For j = 1 To Nsize
            MyBk.Sheets(3).Cells(6, j + 13).Value = InSt.Cells(j, 5).Value & _
                " - " & InSt.Cells(j + 1, 5).Value & " mm"
        Next
        For i = 1 To 26
            MyBk.Sheets(3).Cells(i + 6, 8).Value = InSt.Cells(i, 16).Value
            If Qs(i) > 0 Then _
                MyBk.Sheets(3).Cells(i + 6, 9).Value = Qs(i) * 2650 * 60
            MyBk.Sheets(3).Cells(i + 6, 10).Value = Phi50(i)
            MyBk.Sheets(3).Cells(i + 6, 11).Value = StSt.Cells(i, 10).Value
            If Qs(i) > 0 Then _
                MyBk.Sheets(3).Cells(i + 6, 12).Value = H(i)
            If InSt.Cells(1, 1).Value = "XS" Then
                GetRhAndAreaFromDepth H(i), Rh, Area, Rh1, Area1, Rh2, Area2
                MyBk.Sheets(3).Cells(i + 6, 13).Value = Rh
            End If
            For j = 1 To Nsize
                MyBk.Sheets(3).Cells(i + 6, j + 13).Value = StSt.Cells(i, j + 26).Value * 2650 * 60
            Next
        Next
    End If
    
    'bedload grain size distribution
    If InSt.Cells(6, 1).Value <> "(B)" Then
        MyBk.Sheets(3).Cells(5, 2).Value = "BEDLOAD GRAIN SIZE DISTRIBUTION"
        MyBk.Sheets(3).Cells(6, 2).Value = "Size (mm)"
        MyBk.Sheets(3).Cells(6, 3).Value = "% Finer"
        For i = 1 To Nsize + 1
            MyBk.Sheets(3).Cells(i + 6, 2).Value = InSt.Cells(i, 7).Value
            If InSt.Cells(1, 8).Value < 5 Then 'increasing percent finer
                If i = 1 Then
                    MyBk.Sheets(3).Cells(i + 6, 3).Value = 0
                    dmy = 0
                Else
                    dmy = dmy + p(i - 1) * 100
                    MyBk.Sheets(3).Cells(i + 6, 3).Value = Format(dmy, "###0")
                End If
            Else 'decreasing percent finer
                If i = 1 Then
                    MyBk.Sheets(3).Cells(i + 6, 3).Value = 100
                    dmy = 100
                Else
                    dmy = dmy - p(i - 1) * 100
                    MyBk.Sheets(3).Cells(i + 6, 3).Value = Format(dmy, "###0")
                End If
            End If
        Next
        
        Set MySize = Range(MyBk.Sheets(3).Cells(7, 2), MyBk.Sheets(3).Cells(Nsize + 7, 2))
        Set MyFiner = Range(MyBk.Sheets(3).Cells(7, 3), MyBk.Sheets(3).Cells(Nsize + 7, 3))
        GetGrainSizeStatistics Nsize, MySize, MyFiner, ChD
        
        MyBk.Sheets(3).Cells(Nsize + 10, 2).Value = "STATISTICS OF THE ABOVE GRAIN SIZE DISTRIBUTION:"
        MyBk.Sheets(3).Cells(Nsize + 11, 2).Value = "Geometric mean (mm)"
        MyBk.Sheets(3).Cells(Nsize + 12, 2).Value = "Geometric standard deviation"
        MyBk.Sheets(3).Cells(Nsize + 13, 2).Value = "D10 (mm)"
        MyBk.Sheets(3).Cells(Nsize + 14, 2).Value = "D16 (mm)"
        MyBk.Sheets(3).Cells(Nsize + 15, 2).Value = "D20 (mm)"
        MyBk.Sheets(3).Cells(Nsize + 16, 2).Value = "D50 (mm)"
        MyBk.Sheets(3).Cells(Nsize + 17, 2).Value = "D65 (mm)"
        MyBk.Sheets(3).Cells(Nsize + 18, 2).Value = "D75 (mm)"
        MyBk.Sheets(3).Cells(Nsize + 19, 2).Value = "D84 (mm)"
        MyBk.Sheets(3).Cells(Nsize + 20, 2).Value = "D90 (mm)"
        Range(MyBk.Sheets(3).Cells(Nsize + 10, 2), MyBk.Sheets(3).Cells(Nsize + 20, 2)).HorizontalAlignment = xlGeneral
        For j = 0 To 9
            MyBk.Sheets(3).Cells(Nsize + 11 + j, 4).Value = ChD(j)
        Next
    End If
    
    If InSt.Cells(6, 1).Value = "(A)" Then GoTo 10 'Single discharge
    
    cc = MyBk.Sheets(3).Cells(5, 12).Value & " " & MyBk.Sheets(3).Cells(6, 12).Value

'------
    If InSt.Cells(6, 1).Value = "(A)" Then 'single flow
        'Empty
    ElseIf InSt.Cells(6, 1).Value = "(B)" Then 'between two flows
        'Empty
    Else 'duration curve
        MyBk.Activate
        Worksheets("Input").Select
        Set xRange = Range(MyBk.Worksheets("Input").Cells(5, 5), MyBk.Worksheets("Input").Cells(30, 5))
        Set yRange = Range(MyBk.Worksheets("Input").Cells(5, 6), MyBk.Worksheets("Input").Cells(30, 6))
        AddRatingCurves MyBk, "Input", xRange, yRange, "Plot Duration Curve", _
            "Discharge (cms)", "Exceedance Probability (%)"
        ModifyYaxisToNormal MyBk, "Plot Duration Curve"
        AdjustYaxisScale MyBk, "Plot Duration Curve", 0, 100, 20
    End If
    
    If InSt.Cells(1, 1).Value = "XS" Then 'cross section
        MyBk.Activate
        Worksheets("Input").Select
        Set xRange = Range(MyBk.Worksheets("Input").Cells(16, 2), MyBk.Worksheets("Input").Cells(nXS + 15, 2))
        Set yRange = Range(MyBk.Worksheets("Input").Cells(16, 3), MyBk.Worksheets("Input").Cells(nXS + 15, 3))
        AddRatingCurves MyBk, "Input", xRange, yRange, "Plot XS", "Station (m)", "Elevation (m)"
        ModifyYaxisToNormal MyBk, "Plot XS"
        AdjustXaxisToNormal MyBk, "Plot XS"
    End If
    
    MyBk.Activate
    Worksheets("Input").Select
    Set xRange = Range(MyBk.Worksheets("Input").Cells(4, 8), MyBk.Worksheets("Input").Cells(Nsize + 4, 8))
    Set yRange = Range(MyBk.Worksheets("Input").Cells(4, 9), MyBk.Worksheets("Input").Cells(Nsize + 4, 9))
    AddRatingCurves MyBk, "Input", xRange, yRange, "Plot Sub Size", "Grain Size (mm)", "Percent Finer"
    ModifyYaxisToNormal MyBk, "Plot Sub Size"
    AdjustYaxisScale MyBk, "Plot Sub Size", 0, 100, 20
'------------
    
    MyBk.Activate
    Worksheets("Output").Select
    Set xRange = Range(MyBk.Worksheets("Output").Cells(7, 8), MyBk.Worksheets("Output").Cells(32, 8))
    Set yRange = Range(MyBk.Worksheets("Output").Cells(7, 9), MyBk.Worksheets("Output").Cells(32, 9))
    AddRatingCurves MyBk, "Output", xRange, yRange, "Plot Bedload", _
        "Discharge (cms)", "Bedload Transport Rate (kg/min.)"
    
    If MyOption = 1 Then 'Bakke, add field data to the diagram
        Set xRange = Range(MyBk.Worksheets("Input").Cells(SampleStarts, 8), MyBk.Worksheets("Input").Cells(SampleEnds, 8))
        Set yRange = Range(MyBk.Worksheets("Input").Cells(SampleStarts, 9), MyBk.Worksheets("Input").Cells(SampleEnds, 9))
        AddExperimentalData MyBk, xRange, yRange, "Plot Bedload"
    End If
        
    MyBk.Activate
    Worksheets("Output").Select
    Set xRange = Range(MyBk.Worksheets("Output").Cells(7, 8), MyBk.Worksheets("Output").Cells(32, 8))
    Set yRange = Range(MyBk.Worksheets("Output").Cells(7, 10), MyBk.Worksheets("Output").Cells(32, 10))
    AddRatingCurves MyBk, "Output", xRange, yRange, "Plot Shear", "Discharge (cms)", "Transport Stage" & vbLf & "(Normalized Shields Stress)"
    ModifyYaxisToNormal MyBk, "Plot Shear"
        
    MyBk.Activate
    Worksheets("Output").Select
    Set yRange = Range(MyBk.Worksheets("Output").Cells(7, 12), MyBk.Worksheets("Output").Cells(32, 12))
    AddRatingCurves MyBk, "Output", xRange, yRange, "Plot Depth", "Discharge (cms)", ULCase(cc)
    ModifyYaxisToNormal MyBk, "Plot Depth"
    
10  If MyOption = 0 Then _
        MsgBox "Calculation results with substrate-based bedload equation of Parker and Klingeman (1982) " & _
            "are temporarily stored in workbook " & MyBk.Name & ".  Please save the file with an appropriate " & _
            "file name in an appropriate folder upon finishing of the rest of the run." & vbLf & vbLf & _
            "Click ""OK"" to continue.", vbOKOnly + vbInformation, "Parker and Klingman (1982)"
    If MyOption = 1 Then _
        MsgBox "Calculation results with substrate-based bedload equation of Bakke et al. (1999) " & _
            "are temporarily stored in workbook " & MyBk.Name & ".  Please save the file with an appropriate " & _
            "file name in an appropriate folder upon finishing of the rest of the run." & vbLf & vbLf & _
            "Click ""OK"" to continue.", vbOKOnly + vbInformation, "Bakke et al. (1999)"
End Sub


Sub GoGetTau50StarAndExponent(TauR50Star As Double, Exponent As Double, R As Double, _
    g As Double, Dk As Double, D50mm As Double, Width As Double, f() As Double, _
    p() As Double, Psi() As Double, Slope As Double, Nsize As Long, Nsp As Long)
    
    Dim i As Long, j As Long, MaxIteration As Long
    Dim Rough As Double
    Dim OldTauR50Star As Double, OldExponent As Double
    Dim ExpError As Double, TauError As Double, Tolerance As Double
    
    If PKorPKM = "PK" Then
        TauR50Star = Worksheets("Input").Cells(15, 1).Value
        Exponent = Worksheets("Input").Cells(15, 2).Value
    Else
        TauR50Star = Worksheets("Input").Cells(16, 1).Value
        Exponent = Worksheets("Input").Cells(16, 2).Value
    End If
    i = MsgBox("Optimazation parameters from the last run are:" & vbLf & vbLf & _
        "Taur50 =   " & TauR50Star & vbLf & _
        "Exponent = " & Exponent & vbLf & vbLf & _
        "Apply those parameters? (Click yes to use those parameters and no to start the optimization.)", _
        vbYesNo + vbQuestion, "Optimization Parameters")
    
    If i = vbYes Then Exit Sub
    
    MaxIteration = 6
    
    InitializingProgressBar 1
    
    Tolerance = 0.001
    Rough = Dk * D50mm / 1000
    TauR50Star = 0.0876: Exponent = 0.018  'initial guess set as the default values in Parker and Klingeman (1982)
    ExpError = 1: TauError = 1
    i = 0
    Do While (TauError > Tolerance Or i < 3) And i < MaxIteration
        i = i + 1
        OldTauR50Star = TauR50Star: OldExponent = Exponent
        GoGetTauR50 TauR50Star, Exponent, R, g, Rough, Dk, D50mm, Width, f, Psi, Slope, Nsize, Nsp, i
        GoGetExponent Exponent, TauR50Star, R, g, Dk, D50mm, Width, f, p, Psi, Slope, Nsize, Nsp, i
        ExpError = Abs(Exponent - OldExponent) / Exponent
        If ExpError < Tolerance Then i = 3
        TauError = Abs(TauR50Star - OldTauR50Star) / TauR50Star
        MessageOnWelcome "Calculating with Bakke et al. (1999)" & _
            vbLf & vbLf & "This may take a while!  Please wait..." & _
            vbLf & vbLf & "Finished Iteration No. " & i & " of Maximum " & MaxIteration + 1
        ThisWorkbook.Save
    Loop
    GoGetTauR50 TauR50Star, Exponent, R, g, Rough, Dk, D50mm, Width, f, Psi, Slope, Nsize, Nsp, i
    If PKorPKM = "PK" Then
        Worksheets("Input").Cells(15, 1).Value = TauR50Star
        Worksheets("Input").Cells(15, 2).Value = Exponent
    End If
    If PKorPKM = "PKM" Then
        Worksheets("Input").Cells(16, 1).Value = TauR50Star
        Worksheets("Input").Cells(16, 2).Value = Exponent
    End If
    ThisWorkbook.Save
End Sub

Sub GoGetTauR50(TauR50Star As Double, Exponent As Double, R As Double, g As Double, _
    Rough As Double, Dk As Double, D50mm As Double, Width As Double, f() As Double, _
    Psi() As Double, Slope As Double, Nsize As Long, Nsp As Long, iterationCount As Long)
    
    Dim InSt As Worksheet, StSt As Worksheet
    Dim i As Long, j As Long, kount As Long
    Dim Qw As Double, p(21) As Double, H As Double
    Dim Ustar As Double, Tau50Star As Double
    Dim Rh As Double, Area As Double, Qwc As Double, Qw1 As Double, Qw2 As Double
    Dim lTau As Double, cTau As Double, rTau As Double
    Dim lEr As Double, cEr As Double, rEr As Double, MinEr As Double
    Dim Tolerance As Double, MyError As Double
    Dim TempStor1(1 To 51) As Double, TempStor2(1 To 51) As Double
    
    Set InSt = Worksheets("Input"): Set StSt = Worksheets("Storage")
    
    Tolerance = 0.001
    
    If iterationCount = 1 Then
        lTau = TauR50Star / 5: cTau = TauR50Star: rTau = TauR50Star * 5
    Else
        lTau = TauR50Star / 2: cTau = TauR50Star: rTau = TauR50Star * 2
    End If
    MyError = 1
    Do While MyError > Tolerance
        UpdatingProgressBar 1
        For i = 1 To 51
            If i = 1 Then
                cTau = lTau
            Else
                cTau = cTau * (rTau / lTau) ^ (1 / 50)
            End If
            GetMySquareErrorInTauR _
                Exponent, cTau, R, g, Dk, D50mm, Width, f, p, Psi, Slope, Nsize, Nsp, cEr
            TempStor1(i) = cTau
            TempStor2(i) = cEr
        Next
    
        MinEr = Application.WorksheetFunction.Min(TempStor2)
        For i = 1 To 51
            UpdatingProgressBar 1
            If TempStor2(i) = 0 Then
                Exit For
            ElseIf Abs(TempStor2(i) - MinEr) / TempStor2(i) < 0.00001 Then
                Exit For
            End If
        Next
        cTau = TempStor1(i)
        If i = 1 Then
            lTau = TempStor1(1)
        Else
            lTau = TempStor1(i - 1)
        End If
        If i = 51 Then
            rTau = TempStor1(51)
        Else
            rTau = TempStor1(i + 1)
        End If
        MyError = Abs((rTau - lTau) / cTau)
    Loop
    TauR50Star = cTau
End Sub

Sub GetMySquareErrorInTauR(Exponent As Double, TauR50Star As Double, R As Double, g As Double, _
    Dk As Double, D50mm As Double, Width As Double, f() As Double, p() As Double, _
    Psi() As Double, Slope As Double, Nsize As Long, Nsp As Long, SquareError As Double)

    Dim i As Long, j As Long, k As Long
    Dim InSt As Worksheet, StSt As Worksheet
    Dim Qw As Double, Qs As Double, H As Double, Phi50 As Double
    Dim dmy As Double, Taui As Double
    Dim TempStor1(1 To 200) As Double, TempStor2(1 To 200) As Double
    
    Set InSt = Worksheets("Input"): Set StSt = Worksheets("Storage")
    StSt.Columns("F:G").ClearContents
        
    For i = 1 To Nsp
        UpdatingProgressBar 1
        Qw = InSt.Cells(i, 18).Value
        If InSt.Cells(1, 1).Value = "XS" Then 'cross section
            ParkerklingemanEquationWithCrossSection Exponent, TauR50Star, _
                R, g, Dk, D50mm, Qw, f, p, Psi, Qs, H, Phi50, Slope, Nsize, "Bakke"
        Else ' channel width
            ParkerKlingemanBasicEquation Exponent, TauR50Star, _
                R, g, Dk, D50mm, Width, Qw, f, p, Psi, Qs, H, Phi50, Slope, Nsize, "Bakke"
        End If
        TempStor1(i) = InSt.Cells(i, 19).Value
        TempStor2(i) = Qs * 2650 * 60
    Next
    SquareError = 0
    For i = 1 To Nsp
        If Not InSt.Cells(i, 40).Value And TempStor1(i) < 1E+19 And TempStor2(i) > 0 Then _
            SquareError = SquareError + (Log(TempStor1(i)) - Log(TempStor2(i))) ^ 2
    Next
End Sub

Sub GoGetExponent(Exponent As Double, TauR50Star As Double, R As Double, g As Double, _
    Dk As Double, D50mm As Double, Width As Double, f() As Double, p() As Double, _
    Psi() As Double, Slope As Double, Nsize As Long, Nsp As Long, iterationCount As Long)
    
    Dim StSt As Worksheet
    Dim SquareError As Double
    Dim lExp As Double, cExp As Double, rExp As Double
    Dim lEr As Double, cEr As Double, rEr As Double, MinEr As Double
    Dim i As Long
    Dim Tolerance As Double, MyError As Double
    Dim TempStor1(1 To 51) As Double, TempStor2(1 To 51) As Double

    Tolerance = 0.001
    
    Set StSt = Worksheets("Storage")
    
    If iterationCount = 1 Then
        lExp = Exponent / 5: cExp = Exponent: rExp = cExp * 5
    Else
        lExp = Exponent / 3: cExp = Exponent: rExp = cExp * 3
    End If
    MyError = 1
    Do While MyError > Tolerance
        UpdatingProgressBar 1
        For i = 1 To 51
            If i = 1 Then
                cExp = lExp
            Else
                cExp = cExp * (rExp / lExp) ^ (1 / 50)
            End If
            
            GetMySquareErrorInD50 cExp, TauR50Star, R, g, Dk, D50mm, Width, f, p, Psi, Slope, Nsize, Nsp, cEr
            TempStor1(i) = cExp
            TempStor2(i) = cEr
        Next
        
        MinEr = Application.WorksheetFunction.Min(TempStor2)
        For i = 1 To 51
            If TempStor2(i) = 0 Then
                Exit For
            ElseIf TempStor2(i) - MinEr < 0.00001 Then
                Exit For
            End If
        Next
        
        cExp = TempStor1(i)
        If i = 1 Then
            lExp = TempStor1(i)
        Else
            lExp = TempStor1(i - 1)
        End If
        If i = 51 Then
            rExp = TempStor1(i)
        Else
            rExp = TempStor1(i + 1)
        End If
        MyError = Abs((rExp - lExp) / cExp)
    Loop
    Exponent = cExp

End Sub

Sub GetMySquareErrorInD50(Exponent As Double, TauR50Star As Double, R As Double, g As Double, _
    Dk As Double, D50mm As Double, Width As Double, f() As Double, p() As Double, _
    Psi() As Double, Slope As Double, Nsize As Long, Nsp As Long, SquareError As Double)

    Dim i As Long, j As Long, k As Long
    Dim InSt As Worksheet, StSt As Worksheet
    Dim SizeRange As Range, PctRange As Range
    Dim Qw As Double, Qs As Double, H As Double, Phi50 As Double
    Dim dmy As Double, Di As Double
    
    Set InSt = Worksheets("Input"): Set StSt = Worksheets("Storage")
    Set SizeRange = Range(InSt.Cells(1, 7), InSt.Cells(Nsize + 1, 7))
    Set PctRange = Range(StSt.Cells(1, 7), StSt.Cells(Nsize + 1, 7))
    StSt.Columns("F:G").ClearContents
        
    For i = 1 To Nsp
        UpdatingProgressBar 1
        Qw = InSt.Cells(i, 18).Value
        If InSt.Cells(1, 1).Value = "XS" Then 'cross section
            ParkerklingemanEquationWithCrossSection Exponent, TauR50Star, _
                R, g, Dk, D50mm, Qw, f, p, Psi, Qs, H, Phi50, Slope, Nsize, "Bakke"
        Else ' channel width
            ParkerKlingemanBasicEquation Exponent, TauR50Star, _
                R, g, Dk, D50mm, Width, Qw, f, p, Psi, Qs, H, Phi50, Slope, Nsize, "Bakke"
        End If
        If Qs > 0 Then
            If InSt.Cells(1, 8).Value < 1 Then 'increasing grain size
                dmy = 0
                PctRange.Cells(1).Value = dmy
                For j = 1 To Nsize
                    dmy = dmy + p(j) * 100
                    StSt.Cells(j + 1, 7).Value = Format(dmy, "##0.#")
                Next
                GetCharacteristicGrainSizeinMM PctRange, SizeRange, 50, Di
                StSt.Cells(i, 6).Value = Di
            Else ' decreasing grain size
                dmy = 100
                PctRange.Cells(1).Value = dmy
                For j = 1 To Nsize
                    dmy = dmy - p(j) * 100
                    StSt.Cells(j + 1, 7).Value = Format(dmy, "##0.#")
                Next
                GetCharacteristicGrainSizeinMM PctRange, SizeRange, 50, Di
                StSt.Cells(i, 6).Value = Di
            End If
        Else
            StSt.Cells(i, 6).Value = 1E+20
        End If
    Next
    SquareError = 0
    For i = 1 To Nsp
        If Not InSt.Cells(i, 40).Value And StSt.Cells(i, 5).Value < 1E+19 And StSt.Cells(i, 6).Value < 1E+19 Then _
            SquareError = SquareError + (Log(StSt.Cells(i, 6).Value) - Log(StSt.Cells(i, 5))) ^ 2
    Next
End Sub

Sub AuthorCreateDataForBakkeTesting()
    Dim i As Long, j As Long, InSt As Worksheet
    Dim Qw As Double, Qs As Double, p(21) As Double, f(21) As Double, Psi(21) As Double
    Dim Beta As Double, Taur50 As Double, R As Double, g As Double, Dk As Double, D50mm As Double
    Dim Width As Double, H As Double, Phi50 As Double, Slope As Double, Nsize As Long, dmy As Double
    Dim Pswd As String
    
    Pswd = Application.InputBox("Enter password please:", "Author only")
    
    If Pswd <> "not4you" Then Exit Sub
    
    Set InSt = Worksheets("Input")
    
    InSt.Columns("R:AN").ClearContents
    
    If IsEmpty(InSt.Cells(1, 2)) Then
        Width = InSt.Cells(4, 2).Value - InSt.Cells(4, 1).Value
    Else
        Width = InSt.Cells(1, 2).Value
    End If
    
    Nsize = 0
    Do While Not IsEmpty(InSt.Cells(Nsize + 2, 7))
        Nsize = Nsize + 1
    Loop
    
    For i = 1 To Nsize
        f(i) = Abs(InSt.Cells(i + 1, 8).Value - InSt.Cells(i, 8).Value) / 100
    Next
    For i = 1 To Nsize + 1
        Psi(i) = Log(InSt.Cells(i, 7).Value) / Log(2)
    Next
    
    Beta = 0.018
    Taur50 = 0.0876
    R = 1.65
    g = 9.81
    Dk = 10.7
    D50mm = InSt.Cells(13, 1).Value
    Slope = InSt.Cells(5, 2).Value
    Qw = 45
    InSt.Cells(7, 1).Value = 14
    For i = 1 To InSt.Cells(7, 1).Value
        Qw = Qw * 1.3
        ParkerKlingemanBasicEquation Beta, Taur50, R, g, Dk, D50mm, Width, Qw, f(), _
            p(), Psi(), Qs, H, Phi50, Slope, Nsize, "Bakke"
        Qs = Qs * 2650 * 60 * Exp(2.5 * (Rnd - 0.5))
        InSt.Cells(i, 18).Value = Qw
        InSt.Cells(i, 19).Value = Qs
        For j = 1 To Nsize + 1
            If j = 1 Then
                dmy = 0
            Else
                dmy = dmy + p(j - 1) * 100
            End If
            InSt.Cells(i, j + 20).Value = Format(dmy, "##0.##")
        Next
    Next
End Sub


