Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

' This module implement Parker90 formula

Function BasicGravelTransportRate(R As Double, g As Double, Taursgo As Double, alpha As Double, _
    Beta As Double, Dsg As Double, STD As Double, D90 As Double, Slope As Double, Width As Double, H As Double, _
    Qw As Double, Dk As Double, pp As Range, oo As Range, ss As Range, Size As Long, _
    Psi() As Double, f() As Double, p() As Double, Qs As Double, phisgo As Double) As Integer
    
    Dim i As Long
    Dim Ustar As Double, Rough As Double
    Dim Omega As Double, Omega0 As Double, Sigma0 As Double
    Dim Di As Double
    Dim CheckOut As Label
    Dim nn As Double, nD As Double
    Dim tau As Double, tauT As Double
    Dim Rho As Double
    
    Rho = 1000 'doesn't matter but given a correct value anyway
    
    If OnErrorOn Then On Error GoTo CheckOut
    
' Revised in 2006 to include roughness correction
    Rough = Dk * D90
    
    If Worksheets("Input").Cells(17, 1).Value Then
        nn = Worksheets("Input").Cells(17, 2).Value
        nD = 0.04 * Rough ^ (1 / 6)
        If nD <= nn Then 'apply correction
            H = (nn * Qw / Width / Slope ^ 0.5) ^ (3 / 5)
            tauT = Rho * g * H * Slope
            tau = tauT * (nD / nn) ^ 1.5
            Ustar = (tau / Rho) ^ 0.5
            Worksheets("Input").Cells(17, 2).Interior.ColorIndex = xlNone
        Else 'original method
            H = Width / 100 ' initial guess
            Worksheets("Input").Cells(17, 2).Interior.ColorIndex = 36
            If CalculateShearVelocityWithConstantWidth(g, Qw, Width, H, Slope, Rough, Ustar) = 0 Then
                BasicGravelTransportRate = 0
                Exit Function
            End If
        End If
    Else 'original method
        H = Width / 100 ' initial guess
        If CalculateShearVelocityWithConstantWidth(g, Qw, Width, H, Slope, Rough, Ustar) = 0 Then
            BasicGravelTransportRate = 0
            Exit Function
        End If
    End If
    phisgo = Ustar ^ 2 / R / g / Dsg / Taursgo
    GetOmega0Sigma0 phisgo, Omega0, Sigma0, pp, oo, ss
    Omega = 1 + STD / Sigma0 * (Omega0 - 1)
' 2006 Revision ends here

    Qs = 0
    For i = 1 To Size
        Di = 2 ^ (0.5 * (Psi(i) + Psi(i + 1))) / 1000
        p(i) = GinParker90(Omega * phisgo * (Dsg / Di) ^ Beta) * f(i)
        Qs = Qs + p(i)
    Next
    For i = 1 To Size
        p(i) = p(i) / Qs
    Next
    Qs = alpha * Ustar ^ 3 / R / g * Width * Qs
    BasicGravelTransportRate = 1
    Exit Function
    
CheckOut:
    MsgBox "An error occured while ""BasicGravelTransportRate"" is executed!"
End Function

Function CalculateShearVelocityWithConstantWidth(g As Double, Qw As Double, Width As Double, _
    H As Double, Slope As Double, Rough As Double, Ustar As Double) As Integer
    
    Dim FF As Double, Fp As Double, RE As Double
    Dim Relax As Double
    Dim OldRE As Double
    Dim CheckOut As Label
    
    If OnErrorOn Then On Error GoTo CheckOut
        
    Relax = 1
    RE = 1
    Do While RE > 0.00001
        OldRE = RE
        If H <= 0 Then
            H = Width / ((100 * Rnd) + 5)
            Relax = Relax / 2
            OldRE = 1
        End If
        FF = Qw - 2.5 * Width * H * (g * H * Slope) ^ 0.5 * Log(11 * H / Rough)
        Fp = -2.5 * Width * (g * H * Slope) ^ 0.5 * (1 + 1.5 * Log(11 * H / Rough))
        RE = -FF / Fp
        H = H + Relax * RE
        RE = Abs(RE / H)
        If RE > OldRE Then
            Relax = Relax / 2
        End If
    Loop
    Ustar = (g * H * Slope) ^ 0.5
    CalculateShearVelocityWithConstantWidth = 1
    Exit Function
    
CheckOut:
    MsgBox "An error occured while ""CalculateShearVelocityWithConstantWidth"" is executed!"
End Function

Function CrossSectionalGravelTransportRateWithFloodplain(R As Double, g As Double, Taursgo As Double, alpha As Double, _
    Beta As Double, Dsg As Double, STD As Double, D90 As Double, Slope As Double, Qw As Double, Dk As Double, _
    pp As Range, oo As Range, ss As Range, Size As Long, Psi() As Double, _
    f() As Double, p() As Double, Qs As Double, phisgo As Double, H As Double) As Integer
    
    Dim i As Long
    Dim Ustar As Double, Rh As Double, Area As Double, Rough As Double
    Dim Mn1 As Double, Mn2 As Double ' Manning's n for floodplains
    Dim Rh1 As Double, Rh2 As Double ' hydraulic radius for floodplains
    Dim Area1 As Double, Area2 As Double ' flow area in floodplain
    Dim Qw1 As Double, Qw2 As Double, Qwc As Double ' discharge in floodplains and the main channel
    Dim Omega As Double, Omega0 As Double, Sigma0 As Double
    Dim Di As Double
    Dim CheckOut As Label
    Dim nn As Double, nD As Double
    Dim tau As Double, tauT As Double
    Dim Rho As Double
    
    Rho = 1000 'doesn't matter but given a correct value anyway
    
    If OnErrorOn Then On Error GoTo CheckOut

    Rough = Dk * D90
    Mn1 = Worksheets("Input").Cells(3, 1).Value
    Mn2 = Worksheets("Input").Cells(3, 2).Value
    
' Revised in 2006 to include roughness correction
    If Worksheets("Input").Cells(17, 1).Value Then
        nn = Worksheets("Input").Cells(17, 2).Value
        nD = 0.04 * Rough ^ (1 / 6)
        If nD <= nn Then 'apply correction
            GetDepthWithDischargeManningsn Qw, Slope, nn, Mn1, Mn2, Qwc, Qw1, Qw2, H, Rh, Area
            tauT = Rho * g * Rh * Slope
            tau = tauT * (nD / nn) ^ 1.5
            Ustar = (tau / Rho) ^ 0.5
            Worksheets("Input").Cells(17, 2).Interior.ColorIndex = xlNone
        Else 'original method
            GetDepthWithDischarge Qw, Slope, Rough, Mn1, Mn2, Qwc, Qw1, Qw2, H, Rh, Area, g
            Ustar = (g * Rh * Slope) ^ 0.5
            Worksheets("Input").Cells(17, 2).Interior.ColorIndex = 36
        End If
    Else
        GetDepthWithDischarge Qw, Slope, Rough, Mn1, Mn2, Qwc, Qw1, Qw2, H, Rh, Area, g
        Ustar = (g * Rh * Slope) ^ 0.5
    End If
' 2006 revision ends here

    phisgo = Ustar ^ 2 / R / g / Dsg / Taursgo
    GetOmega0Sigma0 phisgo, Omega0, Sigma0, pp, oo, ss
    Omega = 1 + STD / Sigma0 * (Omega0 - 1)
    
    Qs = 0
    For i = 1 To Size
        Di = 2 ^ (0.5 * (Psi(i) + Psi(i + 1))) / 1000
        p(i) = GinParker90(Omega * phisgo * (Dsg / Di) ^ Beta) * f(i)
        Qs = Qs + p(i)
    Next
    For i = 1 To Size
        p(i) = p(i) / Qs
    Next
    Qs = alpha * Ustar * Slope * Area / R * Qs
    
    CrossSectionalGravelTransportRateWithFloodplain = 1
    Exit Function
    
CheckOut:
    MsgBox "An error occured while ""CrossSectionalGravelTransportRate"" is executed!"
End Function

Sub GetDepthWithDischarge(Qw As Double, Slope As Double, Rough As Double, Mn1 As Double, Mn2 As Double, _
    Qwc As Double, Qw1 As Double, Qw2 As Double, H As Double, Rh As Double, Area As Double, g As Double)
    
    Dim Err As Double 'relative error in iteration
    Dim Hup As Double, Hlw As Double ' upper and lower limit of water depth, with bisect method
    Dim Rh1 As Double, Area1 As Double
    Dim Rh2 As Double, Area2 As Double
    Dim dmy As Double
    
    Hup = Worksheets("Input").Cells(51, 9).Value * 3 'three times of the maximum depth provided in cross section
    Hlw = 0
    H = Hup
    GetRhAndAreaFromDepth H, Rh, Area, Rh1, Area1, Rh2, Area2
    Err = 1
    Do While Err > 0.00001
        If Mn1 > 0.0001 Or Mn1 < 10 Then
            Qw1 = Area1 * Rh1 ^ (2 / 3) * Slope ^ 0.5 / Mn1
        Else
            Qw1 = 0
        End If
        If Mn2 < 0.0001 Or Mn2 < 10 Then
            Qw2 = Area2 * Rh2 ^ (2 / 3) * Slope ^ 0.5 / Mn2
        Else
            Qw2 = 0
        End If
        Qwc = Area * (g * Rh * Slope) ^ 0.5 * 2.5 * Log(11 * Rh / Rough)
        Err = Abs((Qw1 + Qw2 + Qwc - Qw) / Qw)
        If (Qw1 + Qw2 + Qwc) > Qw Then ' decrease H
            Hup = H
            H = 0.5 * (Hup + Hlw)
        ElseIf (Qw1 + Qw2 + Qwc) < Qw Then ' increase H
            Hlw = H
            H = 0.5 * (Hup + Hlw)
        End If
        GetRhAndAreaFromDepth H, Rh, Area, Rh1, Area1, Rh2, Area2
    Loop

End Sub

Function CalculateShearVelocityWithCrossSection(g As Double, Qw As Double, Rh As Double, Area As Double, _
    Slope As Double, Rough As Double, Ustar As Double) As Integer
    
    Dim FF As Double, RE As Double
    Dim Au As Double, Al As Double ' upper and lower bound of area
    Dim Aold As Double
    Dim CheckOut As Label
    
    If OnErrorOn Then On Error GoTo CheckOut

    Area = GetArea(Rh)
    Au = -10
    Al = -10
    
    RE = 1
    Do While RE > 0.00001
        Aold = Area
        FF = Qw / Area / 2.5 / (g * Slope) ^ 0.5 - Rh ^ 0.5 * Log(11 * Rh / Rough)
        If FF > 0 Then ' need to increase area
            Al = Area
            If Au < 0 Then
                Area = Area * 2
            Else
                Area = 0.5 * (Au + Al)
            End If
            RE = Abs(Aold - Area) / Area
        ElseIf FF < 0 Then ' need to decrease area
            Au = Area
            If Al < 0 Then
                Area = Area / 2
            Else
                Area = 0.5 * (Au + Al)
            End If
            RE = Abs(Aold - Area) / Area
        Else ' correct value
            RE = 0
        End If
        Rh = GetRh(Area)
    Loop
    
    Ustar = (g * Rh * Slope) ^ 0.5
    CalculateShearVelocityWithCrossSection = 1
    Exit Function
    
CheckOut:
    MsgBox "An error occured while ""CalculateShearVelocityWithCrossSection"" is executed!"
End Function

Sub GetRhAndAreaFromDepth(H As Double, Rh As Double, Area As Double, Rh1 As Double, Area1 As Double, _
    Rh2 As Double, Area2 As Double)
    Dim i As Integer
    Dim CheckOut As Label
    Dim MySt As Worksheet
    
    Set MySt = Worksheets("Input")
    
    If OnErrorOn Then On Error GoTo CheckOut
    
    If H >= MySt.Cells(51, 9).Value Then
        Area = MySt.Cells(51, 11).Value + (H - MySt.Cells(51, 9).Value) * _
            MySt.Cells(52, 10).Value
        Rh = Area / (MySt.Cells(51, 11).Value / MySt.Cells(51, 10).Value)
        If MySt.Cells(52, 9).Value > 0 And MySt.Cells(51, 12).Value > 0 Then
            Area1 = MySt.Cells(51, 13).Value + (H - MySt.Cells(51, 9).Value) * _
                MySt.Cells(52, 9).Value
            Rh1 = Area1 / (MySt.Cells(51, 13).Value / MySt.Cells(51, 12).Value)
        Else
            Area1 = 0
            Rh1 = 0
        End If
        If MySt.Cells(52, 11).Value > 0 And MySt.Cells(51, 14).Value > 0 Then
            Area2 = MySt.Cells(51, 15).Value + (H - MySt.Cells(51, 9).Value) * _
                MySt.Cells(52, 11).Value
            Rh2 = Area2 / (MySt.Cells(51, 15).Value / MySt.Cells(51, 14).Value)
        Else
            Area2 = 0
            Rh2 = 0
        End If
        Exit Sub
    End If
    
    For i = 2 To 51
        If H >= MySt.Cells(i - 1, 9).Value And H < MySt.Cells(i, 9).Value Then
            Rh = MySt.Cells(i - 1, 10).Value + (MySt.Cells(i, 10).Value - _
                MySt.Cells(i - 1, 10).Value) / (MySt.Cells(i, 9).Value - _
                MySt.Cells(i - 1, 9).Value) * (H - MySt.Cells(i - 1, 9).Value)
            Area = MySt.Cells(i - 1, 11).Value + (MySt.Cells(i, 11).Value - _
                MySt.Cells(i - 1, 11).Value) / (MySt.Cells(i, 9).Value - _
                MySt.Cells(i - 1, 9).Value) * (H - MySt.Cells(i - 1, 9).Value)
        
            If MySt.Cells(52, 9).Value > 0 And MySt.Cells(51, 12).Value > 0 Then
                Rh1 = MySt.Cells(i - 1, 12).Value + (MySt.Cells(i, 12).Value - _
                    MySt.Cells(i - 1, 12).Value) / (MySt.Cells(i, 9).Value - _
                    MySt.Cells(i - 1, 9).Value) * (H - MySt.Cells(i - 1, 9).Value)
                Area1 = MySt.Cells(i - 1, 13).Value + (MySt.Cells(i, 13).Value - _
                    MySt.Cells(i - 1, 13).Value) / (MySt.Cells(i, 9).Value - _
                    MySt.Cells(i - 1, 9).Value) * (H - MySt.Cells(i - 1, 9).Value)
            Else
                Area1 = 0
                Rh1 = 0
            End If
            If MySt.Cells(52, 11).Value > 0 And MySt.Cells(51, 14).Value > 0 Then
                Rh2 = MySt.Cells(i - 1, 14).Value + (MySt.Cells(i, 14).Value - _
                    MySt.Cells(i - 1, 14).Value) / (MySt.Cells(i, 9).Value - _
                    MySt.Cells(i - 1, 9).Value) * (H - MySt.Cells(i - 1, 9).Value)
                Area2 = MySt.Cells(i - 1, 15).Value + (MySt.Cells(i, 15).Value - _
                    MySt.Cells(i - 1, 15).Value) / (MySt.Cells(i, 9).Value - _
                    MySt.Cells(i - 1, 9).Value) * (H - MySt.Cells(i - 1, 9).Value)
            Else
                Area2 = 0
                Rh2 = 0
            End If
            Exit Sub
        End If
    Next
    Exit Sub
    
CheckOut:
    MsgBox "Error when calculating Rh and Area from given depth."
End Sub

Function GetArea(Rh As Double) As Double ' this function is modified from an early model
    Dim i As Integer
    Dim CheckOut As Label
    Dim MySt As Worksheet
    
    If OnErrorOn Then On Error GoTo CheckOut
    
    Set MySt = Worksheets("Input")
    
    If Rh >= MySt.Cells(51, 10).Value Then
        GetArea = MySt.Cells(51, 11).Value + (Rh - MySt.Cells(51, 10).Value) * _
            (MySt.Cells(51, 11).Value / MySt.Cells(51, 10).Value)
        Exit Function
    End If
    
    For i = 2 To 51
        If Rh >= MySt.Cells(i - 1, 10).Value And Rh < MySt.Cells(i, 10).Value Then
            GetArea = MySt.Cells(i - 1, 11).Value + (MySt.Cells(i, 11).Value - _
                MySt.Cells(i - 1, 11).Value) / (MySt.Cells(i, 10).Value - _
                MySt.Cells(i - 1, 10).Value) * (Rh - MySt.Cells(i - 1, 10).Value)
            Exit Function
        End If
    Next
    Exit Function
    
CheckOut:
    MsgBox "An error occured while ""GetArea"" is executed!"
End Function

Function GetRh(Area As Double) As Double ' this function is modified from an early model
    Dim i As Integer
    Dim CheckOut As Label
    Dim MySt As Worksheet
    
    If OnErrorOn Then On Error GoTo CheckOut
    
    Set MySt = Worksheets("Input")
    
    If Area >= MySt.Cells(51, 11).Value Then
        GetRh = MySt.Cells(51, 10).Value + (Area - MySt.Cells(51, 11).Value) / _
            (MySt.Cells(51, 11).Value / MySt.Cells(51, 10).Value)
        Exit Function
    End If
    
    For i = 2 To 51
        If Area >= MySt.Cells(i - 1, 11).Value And Area <= MySt.Cells(i, 11).Value Then
            GetRh = MySt.Cells(i - 1, 10).Value + (MySt.Cells(i, 10).Value - _
                MySt.Cells(i - 1, 10).Value) / (MySt.Cells(i, 11).Value - _
                MySt.Cells(i - 1, 11).Value) * (Area - MySt.Cells(i - 1, 11).Value)
            Exit Function
        End If
    Next
    Exit Function
    
CheckOut:
    MsgBox "An error occured while ""GetRh"" is executed!"
End Function

Sub GetOmega0Sigma0(phisgo As Double, Omega0 As Double, Sigma0 As Double, pp As Range, _
    oo As Range, ss As Range) ' this sub is modified from an early model
    Dim i As Integer
    Dim CheckOut As Label
    
    If OnErrorOn Then On Error GoTo CheckOut
    
    If phisgo <= pp.Cells(1).Value Then
        Omega0 = oo.Cells(1).Value
        Sigma0 = ss.Cells(1).Value
        Exit Sub
    End If
    If phisgo >= pp.Cells(36).Value Then
        Omega0 = oo.Cells(36).Value
        Sigma0 = ss.Cells(36).Value
        Exit Sub
    End If
    
    For i = 2 To 36
        If phisgo <= pp.Cells(i) And phisgo >= pp.Cells(i - 1) Then
            Omega0 = oo.Cells(i - 1) + (oo.Cells(i) - oo.Cells(i - 1)) / _
                (pp.Cells(i) - pp.Cells(i - 1)) * (phisgo - pp.Cells(i - 1))
            Sigma0 = ss.Cells(i - 1) + (ss.Cells(i) - ss.Cells(i - 1)) / _
                (pp.Cells(i) - pp.Cells(i - 1)) * (phisgo - pp.Cells(i - 1))
            Exit For
        End If
    Next
    Exit Sub
    
CheckOut:
    MsgBox "An error occured while ""GetOmega0Sigma0"" is executed!"
End Sub

Function GinParker90(xx As Double) As Double ' this function is copied from an early model without modification
    Dim CheckOut As Label
    
    If OnErrorOn Then On Error GoTo CheckOut
    
    If xx <= 1 Then
        GinParker90 = xx ^ 14.2
    ElseIf xx <= 1.59 Then
        GinParker90 = Exp(14.2 * (xx - 1) - 9.28 * (xx - 1) ^ 2)
    Else
        GinParker90 = 5474 * (1 - 0.853 / xx) ^ 4.5
    End If
    Exit Function
    
CheckOut:
    MsgBox "An error occured while ""GinParker90"" is executed!"
End Function

Sub GetGeometricMeanGrainSizeAndArithmeticStandardDeviation(Nsize As Long, _
    Psi() As Double, f() As Double, Dsg As Double, STD As Double)
    
    Dim i As Long
    
    Dsg = 0
    For i = 1 To Nsize
        Dsg = Dsg + 0.5 * (Psi(i) + Psi(i + 1)) * f(i)
    Next
    STD = 0
    For i = 1 To Nsize
        STD = STD + (0.5 * (Psi(i) + Psi(i + 1)) - Dsg) ^ 2 * f(i)
    Next
    Dsg = 2 ^ Dsg / 1000
    STD = (STD) ^ 0.5

End Sub

Sub PresentResultsForParker90(Qs() As Double, phisgo() As Double, H() As Double, p() As Double)

    Dim i As Long, j As Long, Nsize As Long, dmy As Double, AdjustedNsize As Long
    Dim nXS As Long
    Dim MyBk As Workbook, InSt As Worksheet, StSt As Worksheet
    Dim Rh As Double, Area As Double, Rh1 As Double, Area1 As Double, Rh2 As Double, Area2 As Double
    Dim MySize As Range, MyFiner As Range, ChD(10) As Double
    Dim cc As String
    Dim xRange As Range, yRange As Range
    
    Set InSt = Worksheets("Input")
    Set StSt = Worksheets("Storage")
    
    Nsize = 0
    Do While Not IsEmpty(InSt.Cells(Nsize + 2, 5))
        Nsize = Nsize + 1
    Loop
    
    AdjustedNsize = 0
    If Worksheets("Storage").Cells(1, 1).Value = 0 Then 'adjusted surface grain size exists
        Do While Not IsEmpty(Worksheets("Storage").Cells(AdjustedNsize + 2, 3))
            AdjustedNsize = AdjustedNsize + 1
        Loop
    End If
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
        "Bedload transport equation used: The surface-based bedload equation of Parker (1990)."
    MyBk.Sheets(1).Cells(6, 2).Value = _
        "Input data are stored in worksheet ""Input"" and results are stored in worksheet ""Output""."
    cc = Application.UserName
    If Len(cc) < 1 Or cc = " " Then cc = "unknown"
    MyBk.Sheets(1).Cells(8, 2).Value = _
        "Calculation was performed by " & cc & " on " & Date & "."

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
    If AdjustedNsize > 0 Then 'adjusted surface grain size
        MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 10, 8).Value = "STATISTICS OF THE ABOVE GRAIN SIZE DISTRIBUTIONS:"
        MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 11, 10).Value = "Original"
        MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 11, 11).Value = "Adjusted"
        MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 12, 8).Value = "Geometric mean (mm)"
        MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 13, 8).Value = "Geometric standard deviation"
        MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 14, 8).Value = "D10 (mm)"
        MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 15, 8).Value = "D16 (mm)"
        MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 16, 8).Value = "D25 (mm)"
        MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 17, 8).Value = "D50 (mm)"
        MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 18, 8).Value = "D65 (mm)"
        MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 19, 8).Value = "D75 (mm)"
        MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 20, 8).Value = "D84 (mm)"
        MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 21, 8).Value = "D90 (mm)"
        Range(MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 10, 8), _
            MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 25, 8)).HorizontalAlignment = xlGeneral
        For i = 0 To 9
            MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 12 + i, 10).Value = Format(ChD(i), "###0.##")
        Next
        
        MyBk.Sheets(2).Cells(Nsize + 6, 8).Value = "ADJUSTED SURFACE GRAIN SIZE DISTRIBUTION"
        MyBk.Sheets(2).Cells(Nsize + 6, 8).HorizontalAlignment = xlGeneral
        MyBk.Sheets(2).Cells(Nsize + 7, 8).Value = "Size (mm)"
        MyBk.Sheets(2).Cells(Nsize + 7, 9).Value = "% Finer"
        For i = 1 To AdjustedNsize + 1
            MyBk.Sheets(2).Cells(i + Nsize + 7, 8).Value = _
                Format(Worksheets("Storage").Cells(i, 3).Value, "###0.##")
            MyBk.Sheets(2).Cells(i + Nsize + 7, 9).Value = _
                Format(Worksheets("Storage").Cells(i, 4).Value, "###0.##")
        Next
        Set MySize = Range(Worksheets("Storage").Cells(1, 3), Worksheets("Storage").Cells(AdjustedNsize + 1, 3))
        Set MyFiner = Range(Worksheets("Storage").Cells(1, 4), Worksheets("Storage").Cells(AdjustedNsize + 1, 4))
        GetGrainSizeStatistics AdjustedNsize, MySize, MyFiner, ChD
        For i = 0 To 9
            MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 12 + i, 11).Value = Format(ChD(i), "###0.##")
        Next
        
        If Worksheets("Input").Cells(17, 1).Value Then
            MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 23, 8).Value = "Main channel Manning's n"
            MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 23, 10).Value = _
                Worksheets("Input").Cells(17, 2).Value
            If Worksheets("Input").Cells(17, 2).Interior.ColorIndex <> xlNone Then
                MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 24, 8).Value = "(This main channel Manning's n is not used because it is"
                MyBk.Sheets(2).Cells(Nsize + AdjustedNsize + 25, 8).Value = "smaller than what is calculated based on grain roughness.)"
            End If
        End If
    Else
        MyBk.Sheets(2).Cells(Nsize + 6, 8).Value = "STATISTICS OF THE ABOVE GRAIN SIZE DISTRIBUTION:"
        MyBk.Sheets(2).Cells(Nsize + 7, 8).Value = "Geometric mean (mm)"
        MyBk.Sheets(2).Cells(Nsize + 8, 8).Value = "Geometric standard deviation"
        MyBk.Sheets(2).Cells(Nsize + 9, 8).Value = "D10 (mm)"
        MyBk.Sheets(2).Cells(Nsize + 10, 8).Value = "D16 (mm)"
        MyBk.Sheets(2).Cells(Nsize + 11, 8).Value = "D25 (mm)"
        MyBk.Sheets(2).Cells(Nsize + 12, 8).Value = "D50 (mm)"
        MyBk.Sheets(2).Cells(Nsize + 13, 8).Value = "D65 (mm)"
        MyBk.Sheets(2).Cells(Nsize + 14, 8).Value = "D75 (mm)"
        MyBk.Sheets(2).Cells(Nsize + 15, 8).Value = "D84 (mm)"
        MyBk.Sheets(2).Cells(Nsize + 16, 8).Value = "D90 (mm)"
        Range(MyBk.Sheets(2).Cells(Nsize + 6, 8), MyBk.Sheets(2).Cells(Nsize + 20, 8)).HorizontalAlignment = xlGeneral
        For i = 0 To 9
            MyBk.Sheets(2).Cells(Nsize + 7 + i, 10).Value = Format(ChD(i), "###0.##")
        Next
    
        If Worksheets("Input").Cells(17, 1).Value Then
            MyBk.Sheets(2).Cells(Nsize + 18, 8).Value = "Main channel Manning's n"
            MyBk.Sheets(2).Cells(Nsize + 18, 10).Value = _
                Worksheets("Input").Cells(17, 2).Value
            If Worksheets("Input").Cells(17, 2).Interior.ColorIndex <> xlNone Then
                MyBk.Sheets(2).Cells(Nsize + 19, 8).Value = "(This main channel Manning's n is not used because it is"
                MyBk.Sheets(2).Cells(Nsize + 20, 8).Value = "smaller than what is calculated based on grain roughness.)"
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
    MyBk.Sheets(3).Cells(5, 13).HorizontalAlignment = xlGeneral
    
    ' bedload transport rate
    If InSt.Cells(6, 1).Value = "(A)" Then 'single discharge
        MyBk.Sheets(3).Cells(2, 2).Value = "Bedload transport rate (kg/min.)"
        MyBk.Sheets(3).Cells(2, 6).Value = Qs(1) * 2650 * 60 ' m3/s to kg/min.
        MyBk.Sheets(3).Cells(3, 2).Value = "Normalized Shields stress"
        MyBk.Sheets(3).Cells(3, 2).HorizontalAlignment = xlGeneral
        MyBk.Sheets(3).Cells(3, 6).Value = phisgo(1) ' m3/s to kg/min.
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
        MyBk.Sheets(3).Cells(5, 13).Value = "Bedload transport rate by size, in kg/min."
        If Worksheets("Storage").Cells(1, 1).Value = 0 Then 'adjusted surface grain size exists
            For j = 1 To AdjustedNsize
                MyBk.Sheets(3).Cells(6, j + 12).Value = StSt.Cells(j, 3).Value & _
                    " - " & StSt.Cells(j + 1, 3).Value & " mm"
            Next
        Else
            For j = 1 To Nsize
                MyBk.Sheets(3).Cells(6, j + 12).Value = StSt.Cells(j, 3).Value & _
                    " - " & StSt.Cells(j + 1, 3).Value & " mm"
            Next
        End If
        
        If InSt.Cells(1, 1).Value = "XS" Then
            MyBk.Sheets(3).Cells(5, 11).Value = "Max water"
            MyBk.Sheets(3).Cells(5, 12).Value = "Hydraulic"
            MyBk.Sheets(3).Cells(6, 12).Value = "radius (m)"
        Else
            MyBk.Sheets(3).Cells(5, 11).Value = "Water"
        End If
        MyBk.Sheets(3).Cells(6, 11).Value = "depth (m)"
        For i = 1 To 26
            MyBk.Sheets(3).Cells(i + 6, 8).Value = InSt.Cells(i, 16).Value
            MyBk.Sheets(3).Cells(i + 6, 9).Value = Qs(i) * 2650 * 60
            MyBk.Sheets(3).Cells(i + 6, 10).Value = phisgo(i)
            MyBk.Sheets(3).Cells(i + 6, 11).Value = H(i)
            If InSt.Cells(1, 1).Value = "XS" Then
                GetRhAndAreaFromDepth H(i), Rh, Area, Rh1, Area1, Rh2, Area2
                MyBk.Sheets(3).Cells(i + 6, 12).Value = Rh
            End If
            If Worksheets("Storage").Cells(1, 1).Value = 0 Then 'adjusted surface grain size exists
                For j = 1 To AdjustedNsize
                    MyBk.Sheets(3).Cells(i + 6, j + 12).Value = StSt.Cells(i, j + 26).Value * 2650 * 60
                Next
            Else
                For j = 1 To Nsize
                    MyBk.Sheets(3).Cells(i + 6, j + 12).Value = StSt.Cells(i, j + 26).Value * 2650 * 60
                Next
            End If
        Next
    End If
    
    'bedload grain size distribution
    If InSt.Cells(6, 1).Value <> "(B)" Then
        MyBk.Sheets(3).Cells(5, 2).Value = "BEDLOAD GRAIN SIZE DISTRIBUTION"
        MyBk.Sheets(3).Cells(6, 2).Value = "Size (mm)"
        MyBk.Sheets(3).Cells(6, 3).Value = "% Finer"
        i = 0
        Do While Not IsEmpty(Worksheets("Storage").Cells(i + 1, 3))
            i = i + 1
            MyBk.Sheets(3).Cells(i + 6, 2).Value = Worksheets("Storage").Cells(i, 3).Value
            If Worksheets("Storage").Cells(1, 4).Value < 5 Then 'increasing percent finer
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
    
    If Worksheets("Storage").Cells(1, 1).Value = 0 Then
        MyBk.Sheets(3).Cells(i + 8, 2).HorizontalAlignment = xlGeneral
        MyBk.Sheets(3).Cells(i + 8, 2).Value = "* Fine sediment is excluded from the calculation."
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
    
10  MsgBox "Calculation results with surface-based bedload equation of Parker (1990) " & _
        "are temporarily stored in workbook " & MyBk.Name & ".  Please save the file with an appropriate " & _
        "file name in an appropriate folder upon finishing of the rest of the run." & vbLf & vbLf & _
        "Click ""OK"" to continue.", vbOKOnly + vbInformation, "Parker (1990)"
End Sub



