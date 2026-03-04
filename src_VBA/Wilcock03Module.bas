Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

Function WilcockFunctionG(Phi As Double) As Double
    If Phi > 1.35 Then
        WilcockFunctionG = 14 * (1 - 0.894 / Phi ^ 0.5) ^ 4.5
    Else
        WilcockFunctionG = 0.002 * Phi ^ 7.5
    End If
End Function

Function WilcockTauRi(Di As Double, Dsg As Double, Taursg As Double) As Double
    Dim b As Double ' hiding coefficient
    b = 0.67 / (1 + Exp(1.5 - Di / Dsg))
    WilcockTauRi = (Di / Dsg) ^ b * Taursg
End Function

Function WilcockTauRsg(Rho As Double, R As Double, g As Double, Dsg As Double, Fs As Double) As Double
    WilcockTauRsg = Rho * R * g * Dsg * (0.021 + 0.015 / Exp(20 * Fs))
End Function

Sub BasicWilcock03TransportRate(Nsize As Long, Rho As Double, R As Double, g As Double, _
    Dsg As Double, Qw As Double, Taursg As Double, Rough As Double, Slope As Double, _
    Psi() As Double, f() As Double, p() As Double, Qs As Double, Phi As Double, H As Double)
    
    Dim InSt As Worksheet
    Dim Width As Double
    Dim Ustar As Double, tau As Double, Di As Double, MyPhi As Double
    Dim nn As Double 'Manning's n
    Dim nD As Double 'Manning's n associated with particle grain
    Dim tauT As Double 'overall shear stress
    Dim i As Long
    
    
    Set InSt = Worksheets("Input")
    
    Width = InSt.Cells(1, 2).Value
    
' Revised in 2006 to include roughness correction
    If Worksheets("Input").Cells(20, 1).Value Then 'Manning's n correction
        nn = Worksheets("Input").Cells(20, 2).Value
        nD = 0.04 * Rough ^ (1 / 6)
        If nD <= nn Then 'apply correction
            H = (nn * Qw / Width / Slope ^ 0.5) ^ (3 / 5)
            tauT = Rho * g * H * Slope
            tau = tauT * (nD / nn) ^ 1.5
            Ustar = (tau / Rho) ^ 0.5
            Worksheets("Input").Cells(20, 2).Interior.ColorIndex = xlNone
        Else 'original equation
            Worksheets("Input").Cells(20, 2).Interior.ColorIndex = 36
            H = (Qw / 8.1 / Width) ^ 0.6 * Rough ^ 0.1 / (g * Slope) ^ 0.3
            Ustar = (g * H * Slope) ^ 0.5
            tau = Rho * Ustar ^ 2
        End If
    Else 'original equation
        H = (Qw / 8.1 / Width) ^ 0.6 * Rough ^ 0.1 / (g * Slope) ^ 0.3
        Ustar = (g * H * Slope) ^ 0.5
        tau = Rho * Ustar ^ 2
    End If
    Phi = tau / Taursg
' 2006 revision ends here
    
    Qs = 0
    For i = 1 To Nsize
        Di = 2 ^ (0.5 * (Psi(i) + Psi(i + 1))) / 1000
        MyPhi = tau / WilcockTauRi(Di, Dsg, Taursg)
        p(i) = WilcockFunctionG(MyPhi) * f(i)
        Qs = Qs + p(i)
    Next
    If Qs > 0 Then
        For i = 1 To Nsize
            p(i) = p(i) / Qs
        Next
    End If
    Qs = Qs * Width * Ustar ^ 3 / R / g
End Sub

Sub CrossSectionWilcock03TransportRate(Nsize As Long, Rho As Double, R As Double, g As Double, _
    Dsg As Double, Qw As Double, Taursg As Double, Rough As Double, Slope As Double, _
    Psi() As Double, f() As Double, p() As Double, Qs As Double, Phi As Double, H As Double)
    
    Dim InSt As Worksheet
    Dim Width As Double
    Dim Ustar As Double, tau As Double, Di As Double, MyPhi As Double
    Dim Qwc As Double, Qw1 As Double, Qw2 As Double, Rh As Double, Area As Double
    Dim nn As Double 'Manning's n
    Dim nD As Double 'Manning's n associated with particle grain
    Dim tauT As Double 'overall shear stress
    Dim i As Long

    Set InSt = Worksheets("Input")

' Revised in 2006 to include roughness correction
    If Worksheets("Input").Cells(20, 1).Value Then 'Manning's n correction
        nn = Worksheets("Input").Cells(20, 2).Value
        nD = 0.04 * Rough ^ (1 / 6)
        If nD <= nn Then 'apply correction
            GetDepthWithDischargeManningsn Qw, Slope, nn, InSt.Cells(3, 1).Value, _
                InSt.Cells(3, 2).Value, Qwc, Qw1, Qw2, H, Rh, Area
            tauT = Rho * g * Rh * Slope
            tau = tauT * (nD / nn) ^ 1.5
            Ustar = (tau / Rho) ^ 0.5
            Worksheets("Input").Cells(20, 2).Interior.ColorIndex = xlNone
        Else 'original method
            Worksheets("Input").Cells(20, 2).Interior.ColorIndex = 36
            GetDepthWithDischargeManningStrickler Qw, Slope, Rough, InSt.Cells(3, 1).Value, _
                InSt.Cells(3, 2).Value, Qwc, Qw1, Qw2, H, Rh, Area, g
            Ustar = (g * Rh * Slope) ^ 0.5
            tau = Rho * Ustar ^ 2
        End If
    Else 'original method
        GetDepthWithDischargeManningStrickler Qw, Slope, Rough, InSt.Cells(3, 1).Value, _
            InSt.Cells(3, 2).Value, Qwc, Qw1, Qw2, H, Rh, Area, g
        Ustar = (g * Rh * Slope) ^ 0.5
        tau = Rho * Ustar ^ 2
    End If
    Phi = tau / Taursg
' 2006 revision ends here
    
    Qs = 0
    For i = 1 To Nsize
        Di = 2 ^ (0.5 * (Psi(i) + Psi(i + 1))) / 1000
        MyPhi = tau / WilcockTauRi(Di, Dsg, Taursg)
        p(i) = WilcockFunctionG(MyPhi) * f(i)
        Qs = Qs + p(i)
    Next
    If Qs > 0 Then
        For i = 1 To Nsize
            p(i) = p(i) / Qs
        Next
    End If
    Qs = Qs * Area * Slope * Ustar / R
End Sub

Sub PresentResultsForWilcock03Results(Qs() As Double, Phi() As Double, H() As Double, p() As Double)
    Dim i As Long, j As Long, Nsize As Long, dmy As Double
    Dim nXS As Long
    Dim MyBk As Workbook, InSt As Worksheet
    Dim Rh As Double, Area As Double, Rh1 As Double, Area1 As Double, Rh2 As Double, Area2 As Double
    Dim MySize As Range, MyFiner As Range, ChD(10) As Double
    Dim xRange As Range, yRange As Range
    Dim cc As String
    
    Set InSt = Worksheets("Input")
    
    Nsize = 0
    Do While Not IsEmpty(InSt.Cells(Nsize + 2, 5))
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
    MyBk.Sheets(1).Cells(4, 2).Value = _
        "Bedload transport equation used: The surface-based bedload equation of Wilcock and Crowe (2003)."
    MyBk.Sheets(1).Cells(6, 2).Value = _
        "Input data are stored in worksheet ""Input"" and results are stored in worksheet ""Output""."
    MyBk.Sheets(1).Cells(8, 2).Value = _
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
    
    ' input date: surface grain size distribution
    MyBk.Sheets(2).Cells(2, 8).Value = "SURFACE GRAIN SIZE DISTRIBUTION"
    MyBk.Sheets(2).Cells(3, 8).Value = "Size (mm)"
    MyBk.Sheets(2).Cells(3, 9).Value = "% Finer"
    For i = 1 To Nsize + 1
        MyBk.Sheets(2).Cells(i + 3, 8).Value = Format(InSt.Cells(i, 5).Value, "###0.##")
        MyBk.Sheets(2).Cells(i + 3, 9).Value = Format(InSt.Cells(i, 6).Value, "###0.##")
    Next
    
    Set MySize = Range(InSt.Cells(1, 5), InSt.Cells(Nsize + 1, 5))
    Set MyFiner = Range(InSt.Cells(1, 6), InSt.Cells(Nsize + 1, 6))
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
        Range(MyBk.Sheets(2).Cells(Nsize + 7, 8), MyBk.Sheets(2).Cells(Nsize + 21, 8)).HorizontalAlignment = xlGeneral
        For i = 0 To 9
            MyBk.Sheets(2).Cells(Nsize + 8 + i, 10).Value = ChD(i)
        Next
    
        If Worksheets("Input").Cells(20, 1).Value Then 'Manning's n correction
            MyBk.Sheets(2).Cells(Nsize + 19, 8).Value = "Main channel Manning's n"
            MyBk.Sheets(2).Cells(Nsize + 19, 10).Value = _
                Worksheets("Input").Cells(20, 2).Value
            If Worksheets("Input").Cells(20, 2).Interior.ColorIndex <> xlNone Then
                MyBk.Sheets(2).Cells(Nsize + 20, 8).Value = "(This main channel Manning's n is not used because it is"
                MyBk.Sheets(2).Cells(Nsize + 21, 8).Value = "smaller than what is calculated based on grain roughness.)"
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
    MyBk.Sheets(3).Cells(5, 13).HorizontalAlignment = xlGeneral
    
    ' bedload transport rate
    If InSt.Cells(6, 1).Value = "(A)" Then 'single discharge
        MyBk.Sheets(3).Cells(2, 2).Value = "Bedload transport rate (kg/min.)"
        MyBk.Sheets(3).Cells(2, 6).Value = Qs(1) * 2650 * 60 ' m3/s to kg/min.
        MyBk.Sheets(3).Cells(3, 2).Value = "Normalized Shields stress"
        MyBk.Sheets(3).Cells(3, 2).HorizontalAlignment = xlGeneral
        MyBk.Sheets(3).Cells(3, 6).Value = Phi(1) ' m3/s to kg/min.
    Else
        If InSt.Cells(6, 1).Value = "(B)" Then 'min. and max. discharge
            MyBk.Sheets(3).Cells(2, 2).Value = "Rating curves are presented starting Column H."
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
        If InSt.Cells(1, 1).Value = "XS" Then
            MyBk.Sheets(3).Cells(5, 11).Value = "Max water"
            MyBk.Sheets(3).Cells(5, 12).Value = "Hydraulic"
            MyBk.Sheets(3).Cells(6, 12).Value = "radius (m)"
        Else
            MyBk.Sheets(3).Cells(5, 11).Value = "Water"
        End If
        MyBk.Sheets(3).Cells(6, 11).Value = "depth (m)"
        MyBk.Sheets(3).Cells(5, 13).Value = "Sediment transport rate by size, in kg/min."
        For j = 1 To Nsize
            MyBk.Sheets(3).Cells(6, j + 12).Value = InSt.Cells(j, 5).Value & _
                " - " & InSt.Cells(j + 1, 5).Value & " mm"
        Next
        For i = 1 To 26
            MyBk.Sheets(3).Cells(i + 6, 8).Value = InSt.Cells(i, 16).Value
            If Qs(i) > 0 Then _
                MyBk.Sheets(3).Cells(i + 6, 9).Value = Qs(i) * 2650 * 60
            MyBk.Sheets(3).Cells(i + 6, 10).Value = Phi(i)
            MyBk.Sheets(3).Cells(i + 6, 11).Value = H(i)
            If InSt.Cells(1, 1).Value = "XS" Then
                GetRhAndAreaFromDepth H(i), Rh, Area, Rh1, Area1, Rh2, Area2
                MyBk.Sheets(3).Cells(i + 6, 12).Value = Rh
            End If
            For j = 1 To Nsize
                MyBk.Sheets(3).Cells(i + 6, j + 12).Value = Worksheets("Storage").Cells(i, j + 26).Value * 2650 * 60
            Next
        Next
    End If
    
    'bedload grain size distribution
    If InSt.Cells(6, 1).Value <> "(B)" Then
        MyBk.Sheets(3).Cells(5, 2).Value = "BEDLOAD GRAIN SIZE DISTRIBUTION"
        MyBk.Sheets(3).Cells(6, 2).Value = "Size (mm)"
        MyBk.Sheets(3).Cells(6, 3).Value = "% Finer"
        i = 0
        Do While Not IsEmpty(InSt.Cells(i + 1, 5))
            i = i + 1
            MyBk.Sheets(3).Cells(i + 6, 2).Value = InSt.Cells(i, 5).Value
            If InSt.Cells(1, 6).Value < 5 Then 'increasing percent finer
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
        Loop
        i = i - 1
        Set MySize = Range(MyBk.Sheets(3).Cells(7, 2), MyBk.Sheets(3).Cells(i + 7, 2))
        Set MyFiner = Range(MyBk.Sheets(3).Cells(7, 3), MyBk.Sheets(3).Cells(i + 7, 3))
        GetGrainSizeStatistics i, MySize, MyFiner, ChD
        
        MyBk.Sheets(3).Cells(i + 10, 2).Value = "STATISTICS OF THE ABOVE GRAIN SIZE DISTRIBUTION:"
        MyBk.Sheets(3).Cells(i + 11, 2).Value = "Geometric mean (mm)"
        MyBk.Sheets(3).Cells(i + 12, 2).Value = "Geometric standard deviation"
        MyBk.Sheets(3).Cells(i + 13, 2).Value = "D10 (mm)"
        MyBk.Sheets(3).Cells(i + 14, 2).Value = "D16 (mm)"
        MyBk.Sheets(3).Cells(i + 15, 2).Value = "D25 (mm)"
        MyBk.Sheets(3).Cells(i + 16, 2).Value = "D50 (mm)"
        MyBk.Sheets(3).Cells(i + 17, 2).Value = "D65 (mm)"
        MyBk.Sheets(3).Cells(i + 18, 2).Value = "D75 (mm)"
        MyBk.Sheets(3).Cells(i + 19, 2).Value = "D84 (mm)"
        MyBk.Sheets(3).Cells(i + 20, 2).Value = "D90 (mm)"
        Range(MyBk.Sheets(3).Cells(i + 10, 2), MyBk.Sheets(3).Cells(i + 20, 2)).HorizontalAlignment = xlGeneral
        For j = 0 To 9
            MyBk.Sheets(3).Cells(i + 11 + j, 4).Value = Format(ChD(j), "###0.##")
        Next
    End If
    
    If InSt.Cells(6, 1).Value = "(A)" Then GoTo 10 'Single discharge
    
    cc = MyBk.Sheets(3).Cells(5, 11).Value & " " & MyBk.Sheets(3).Cells(6, 11).Value
    
'-------
    If InSt.Cells(6, 1).Value = "(A)" Then 'single flow
        'Empty
    ElseIf InSt.Cells(6, 1).Value = "(B)" Then 'between two flows
        'Empty
    Else 'flow duration curve
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
    AddRatingCurves MyBk, "Input", xRange, yRange, "Plot Surf Size", "Grain Size (mm)", "Percent Finer"
    ModifyYaxisToNormal MyBk, "Plot Surf Size"
    AdjustYaxisScale MyBk, "Plot Surf Size", 0, 100, 20
'---------

    MyBk.Activate
    Worksheets("Output").Select
    Set xRange = Range(MyBk.Worksheets("Output").Cells(7, 8), MyBk.Worksheets("Output").Cells(32, 8))
    Set yRange = Range(MyBk.Worksheets("Output").Cells(7, 9), MyBk.Worksheets("Output").Cells(32, 9))
    AddRatingCurves MyBk, "Output", xRange, yRange, "Plot Bedload", _
        "Discharge (cms)", "Bedload Transport Rate (kg/min.)"
    
    MyBk.Activate
    Worksheets("Output").Select
    Set yRange = Range(MyBk.Worksheets("Output").Cells(7, 10), MyBk.Worksheets("Output").Cells(32, 10))
    AddRatingCurves MyBk, "Output", xRange, yRange, "Plot Shear", "Discharge (cms)", "Transport Stage" & vbLf & "(Normalized Shields Stress)"
    ModifyYaxisToNormal MyBk, "Plot Shear"
    
    MyBk.Activate
    Worksheets("Output").Select
    Set yRange = Range(MyBk.Worksheets("Output").Cells(7, 11), MyBk.Worksheets("Output").Cells(32, 11))
    AddRatingCurves MyBk, "Output", xRange, yRange, "Plot Depth", "Discharge (cms)", ULCase(cc)
    ModifyYaxisToNormal MyBk, "Plot Depth"
    
10  MsgBox "Calculation results with surface-based bedload equation of Wilcock and Crowe (2003) " & _
        "are temporarily stored in workbook " & MyBk.Name & ".  Please save the file with an appropriate " & _
        "file name in an appropriate folder upon finishing of the rest of the run." & vbLf & vbLf & _
        "Click ""OK"" to continue.", vbOKOnly + vbInformation, "Wilcock and Crowe (2003)"
End Sub


