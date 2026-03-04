Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

' This module stores numerical procesures for bedload equations

Sub SurfaceBasedParker90(dmy As Integer)
    Dim InSt As Worksheet, CpSt As Worksheet, StSt As Worksheet
    Dim Qs(27) As Double, phisgo(27) As Double
    Dim ppRange As Range, ooRange As Range, ssRange As Range
    Dim fRange As Range, dRange As Range
    Dim Dsg As Double, STD As Double, H(26) As Double, D90 As Double
    Dim Psi(21) As Double, f(21) As Double, p(21) As Double, Ap(21) As Double
    Dim Nsize As Long 'number of grain size fractions
    Dim Ncal As Long ' number of calculations
    Dim i As Long, j As Long, k As Long
    Dim Sum As Double, Dura As Double
    
    If Parker90 Then
        MessageOnWelcome "Calculating with the surface-based bedload equation of " & _
            "Parker (1990)." & vbLf & vbLf & "Please wait..."
            
        Set InSt = Worksheets("Input")
        Set CpSt = Worksheets("cp")
        Set StSt = Worksheets("Storage")
        
        Set ppRange = Range(CpSt.Cells(2, 1), CpSt.Cells(37, 1))
        Set ooRange = Range(CpSt.Cells(2, 2), CpSt.Cells(37, 2))
        Set ssRange = Range(CpSt.Cells(2, 3), CpSt.Cells(37, 3))
        
        Nsize = 0
        Do While Not IsEmpty(Worksheets("Storage").Cells(Nsize + 1, 3))
            Nsize = Nsize + 1
        Loop
        Nsize = Nsize - 1
        
        Set dRange = Range(StSt.Cells(1, 3), StSt.Cells(Nsize + 1, 3))
        Set fRange = Range(StSt.Cells(1, 4), StSt.Cells(Nsize + 1, 4))
        
        For i = 1 To Nsize + 1
            Psi(i) = Log(Worksheets("Storage").Cells(i, 3).Value) / Log(2)
        Next
        For i = 1 To Nsize
            Ap(i) = 0
            f(i) = Abs(Worksheets("Storage").Cells(i + 1, 4).Value - _
                Worksheets("Storage").Cells(i, 4).Value) / 100
        Next
        
        GetGeometricMeanGrainSizeAndArithmeticStandardDeviation Nsize, Psi, f, Dsg, STD
        
        GetCharacteristicGrainSizeinMM fRange, dRange, 90, D90
        D90 = D90 / 1000 ' mm ==> m
    
        If InSt.Cells(6, 1).Value = "(A)" Then 'single discharge
            Ncal = 1
        Else ' multiple (26) discharges
            Ncal = 26
        End If
        If InSt.Cells(1, 1).Value = "XS" Then 'cross section is used
            For i = 1 To Ncal
                k = CrossSectionalGravelTransportRateWithFloodplain(1.65, 9.81, 0.0386, 0.00218, _
                    0.0951, Dsg, STD, D90, InSt.Cells(5, 2).Value, InSt.Cells(i, 16).Value, 2, _
                    ppRange, ooRange, ssRange, Nsize, Psi, f, p, Qs(i), phisgo(i), H(i))
                For k = 1 To Nsize
                    Worksheets("Storage").Cells(i, k + 26).Value = Qs(i) * p(k)
                Next
                If Left(InSt.Cells(6, 1).Value, 2) = "(C" Then
                    If i = 1 Then
                        Dura = 0.5 * (InSt.Cells(1, 17).Value - InSt.Cells(2, 17).Value) / 100
                    ElseIf i = 26 Then
                        Dura = 0.5 * (InSt.Cells(25, 17).Value - InSt.Cells(26, 17).Value) / 100
                    Else
                        Dura = 0.5 * (InSt.Cells(i - 1, 17).Value - InSt.Cells(i + 1, 17).Value) / 100
                    End If
                    For j = 1 To Nsize
                        Ap(j) = Ap(j) + Qs(i) * p(j) * Dura
                    Next
                End If
            Next 'i=1 to Ncal
        Else ' bankfull width is used
            For i = 1 To Ncal
                k = BasicGravelTransportRate(1.65, 9.81, 0.0386, 0.00218, 0.0951, _
                    Dsg, STD, D90, InSt.Cells(5, 2).Value, InSt.Cells(1, 2).Value, H(i), _
                    InSt.Cells(i, 16).Value, 2, ppRange, ooRange, ssRange, Nsize, _
                    Psi, f, p, Qs(i), phisgo(i))
                For k = 1 To Nsize
                    Worksheets("Storage").Cells(i, k + 26).Value = Qs(i) * p(k)
                Next
                If Left(InSt.Cells(6, 1).Value, 2) = "(C" Then
                    If i = 1 Then
                        Dura = 0.5 * (InSt.Cells(1, 17).Value - InSt.Cells(2, 17).Value) / 100
                    ElseIf i = 26 Then
                        Dura = 0.5 * (InSt.Cells(25, 17).Value - InSt.Cells(26, 17).Value) / 100
                    Else
                        Dura = 0.5 * (InSt.Cells(i - 1, 17).Value - InSt.Cells(i + 1, 17).Value) / 100
                    End If
                    For j = 1 To Nsize
                        Ap(j) = Ap(j) + Qs(i) * p(j) * Dura
                    Next
                End If
            Next 'i=1 to Ncal
        End If
        If Left(InSt.Cells(6, 1).Value, 2) = "(C" Then
            Sum = 0
            For j = 1 To Nsize
                Sum = Sum + Ap(j)
            Next
            If Sum > 0 Then
                For j = 1 To Nsize
                    Ap(j) = Ap(j) / Sum
                Next
            End If
        End If
        If InSt.Cells(6, 1).Value = "(A)" Then
            For j = 1 To Nsize
                Ap(j) = p(j)
            Next
        End If
        EndMessageOnWelcome 1
        PresentResultsForParker90 Qs, phisgo, H, Ap
        ModifyMenu
    End If
    SubstrateBasedParker82 1
End Sub

Sub SubstrateBasedParker82(dmy As Integer)
    Dim Ncal As Long, i As Long, j As Long, k As Long
    Dim InSt As Worksheet
    Dim Qs(27) As Double, H(27) As Double, Phi50(27) As Double
    
    Set InSt = Worksheets("Input")
    
    If Parker82 Then
        MessageOnWelcome "Calculating with Parker-Klingeman-McLean (D50-based, 1982)." & _
            vbLf & vbLf & "Please wait..."
        If InSt.Cells(6, 1).Value = "(A)" Then 'single discharge
            Ncal = 1
        Else ' multiple (26) discharges
            Ncal = 26
        End If
        If InSt.Cells(1, 1).Value = "XS" Then 'cross section is used
            For i = 1 To Ncal
                k = Parker82CrossSectionWithFloodplains(InSt.Cells(i, 16).Value, _
                    InSt.Cells(13, 1).Value, InSt.Cells(5, 2).Value, Qs(i), H(i), Phi50(i), _
                    1.65, 9.81, 0.0025, 0.0876, 10.7)
            Next 'i=1 to Ncal
        Else ' bankfull width is used
            For i = 1 To Ncal
                k = BasicParker82(InSt.Cells(i, 16).Value, InSt.Cells(13, 1).Value, _
                    InSt.Cells(5, 2).Value, InSt.Cells(1, 2).Value, Qs(i), H(i), Phi50(i), _
                    1.65, 9.81, 0.0025, 0.0876, 10.7)
            Next 'i=1 to Ncal
        End If
        EndMessageOnWelcome 1
        PresentResultsForParker82 Qs, Phi50, H
        ModifyMenu
    End If
    ParkerKlingeman82 1
End Sub

Function ParkerKlingeman82(dmy As Integer) As Integer 'Parker-Klingeman 1982
    Dim Ncal As Long, Nsize As Long, i As Long, j As Long, k As Integer
    Dim InSt As Worksheet, StSt As Worksheet
    Dim Psi(21) As Double, p(21) As Double, f(21) As Double, Ap(21) As Double
    Dim Qs(27) As Double, H(27) As Double, Phi50(27) As Double
    Dim Sum As Double, Dura As Double
    Dim SizeRange As Range, PctRange As Range, D50 As Double
        
    If PK82 Then
        PKorPKM = "PK"
        MessageOnWelcome "Calculating with Parker-Klingeman (substrate-based, 1982)." & _
            vbLf & vbLf & "Please wait..."
        
        Set InSt = Worksheets("Input"): Set StSt = Worksheets("Storage")
        
        Nsize = 0
        Do While Not IsEmpty(InSt.Cells(Nsize + 2, 7))
            Nsize = Nsize + 1
        Loop
        
        Set SizeRange = Range(InSt.Cells(1, 7), InSt.Cells(Nsize + 2, 7))
        Set PctRange = Range(StSt.Cells(1, 7), StSt.Cells(Nsize + 2, 7))
        
        For i = 1 To Nsize + 1
            Psi(i) = Log(InSt.Cells(i, 7).Value) / Log(2)
        Next
        For i = 1 To Nsize
            Ap(i) = 0
            f(i) = Abs(InSt.Cells(i + 1, 8).Value - InSt.Cells(i, 8).Value) / 100
        Next
    
        If InSt.Cells(6, 1).Value = "(A)" Then 'single discharge
            Ncal = 1
        Else
            Ncal = 26
        End If
        If InSt.Cells(1, 1).Value = "XS" Then 'cross section is used
            For i = 1 To Ncal
                k = ParkerklingemanEquationWithCrossSection(0.018, 0.0876, 1.65, 9.81, _
                    10.7, InSt.Cells(13, 1).Value, InSt.Cells(i, 16).Value, f, p, Psi, Qs(i), H(i), Phi50(i), _
                    InSt.Cells(5, 2).Value, Nsize, "PK")
                For k = 1 To Nsize
                    StSt.Cells(i, k + 26).Value = Qs(i) * p(k)
                Next
                If Left(InSt.Cells(6, 1).Value, 2) = "(C" Then
                    If i = 1 Then
                        Dura = 0.5 * (InSt.Cells(1, 17).Value - InSt.Cells(2, 17).Value) / 100
                    ElseIf i = 26 Then
                        Dura = 0.5 * (InSt.Cells(25, 17).Value - InSt.Cells(26, 17).Value) / 100
                    Else
                        Dura = 0.5 * (InSt.Cells(i - 1, 17).Value - InSt.Cells(i + 1, 17).Value) / 100
                    End If
                    For j = 1 To Nsize
                        Ap(j) = Ap(j) + Qs(i) * p(j) * Dura
                    Next
                End If
                
                If InSt.Cells(1, 8).Value < 1 Then 'increasing size
                    D50 = 0
                    StSt.Cells(1, 7).Value = D50
                    For j = 1 To Nsize
                        D50 = D50 + p(j) * 100
                        StSt.Cells(j + 1, 7).Value = Format(D50, "##0.#")
                    Next
                Else 'decreasing size
                    D50 = 100
                    StSt.Cells(1, 7).Value = D50
                    For j = 1 To Nsize
                        D50 = D50 - p(j) * 100
                        StSt.Cells(j + 1, 7).Value = Format(D50, "##0.#")
                    Next
                End If
                GetCharacteristicGrainSizeinMM PctRange, SizeRange, 50, D50
                StSt.Cells(i, 10).Value = Format(D50, "##0.#")
            Next 'i=1 to Ncal
        Else 'bankfull width is used
            For i = 1 To Ncal
                k = ParkerKlingemanBasicEquation(0.018, 0.0876, 1.65, 9.81, 10.7, _
                    InSt.Cells(13, 1).Value, InSt.Cells(1, 2).Value, InSt.Cells(i, 16).Value, _
                    f, p, Psi, Qs(i), H(i), Phi50(i), InSt.Cells(5, 2).Value, Nsize, "PK")
                For k = 1 To Nsize
                    StSt.Cells(i, k + 26).Value = Qs(i) * p(k)
                Next
                If Left(InSt.Cells(6, 1).Value, 2) = "(C" Then
                    If i = 1 Then
                        Dura = 0.5 * (InSt.Cells(1, 17).Value - InSt.Cells(2, 17).Value) / 100
                    ElseIf i = 26 Then
                        Dura = 0.5 * (InSt.Cells(25, 17).Value - InSt.Cells(26, 17).Value) / 100
                    Else
                        Dura = 0.5 * (InSt.Cells(i - 1, 17).Value - InSt.Cells(i + 1, 17).Value) / 100
                    End If
                    For j = 1 To Nsize
                        Ap(j) = Ap(j) + Qs(i) * p(j) * Dura
                    Next
                End If
                
                If InSt.Cells(1, 8).Value < 1 Then 'increasing size
                    D50 = 0
                    StSt.Cells(1, 7).Value = D50
                    For j = 1 To Nsize
                        D50 = D50 + p(j) * 100
                        StSt.Cells(j + 1, 7).Value = Format(D50, "##0.#")
                    Next
                Else 'decreasing size
                    D50 = 100
                    StSt.Cells(1, 7).Value = D50
                    For j = 1 To Nsize
                        D50 = D50 - p(j) * 100
                        StSt.Cells(j + 1, 7).Value = Format(D50, "##0.#")
                    Next
                End If
                GetCharacteristicGrainSizeinMM PctRange, SizeRange, 50, D50
                StSt.Cells(i, 10).Value = Format(D50, "##0.#")
            Next
        End If
        If Left(InSt.Cells(6, 1).Value, 2) = "(C" Then
            Sum = 0
            For j = 1 To Nsize
                Sum = Sum + Ap(j)
            Next
            If Sum > 0 Then
                For j = 1 To Nsize
                    Ap(j) = Ap(j) / Sum
                Next
            End If
        End If
        If InSt.Cells(6, 1).Value = "(A)" Then
            For j = 1 To Nsize
                Ap(j) = p(j)
            Next
        End If
        EndMessageOnWelcome 1
        PresentResultsForParkerKlingeman82 Qs, Phi50, H, Ap, 0
        ModifyMenu
    End If
    WilcockTwoFractionModel 1
End Function

Sub WilcockTwoFractionModel(dmy As Integer)
    Dim Ncal As Long, Nsp As Long, i As Long, k As Long
    Dim InSt As Worksheet
    Dim Qs(27) As Double, QG(27) As Double, H(27) As Double
    Dim Qw As Double, Qw1 As Double, Qw2 As Double, Qwc As Double
    Dim Width As Double, Slope As Double, Area As Double, Rh As Double
    Dim Rough As Double, TaurG As Double, TaurS As Double
    Dim Ustar As Double, tau As Double
    Dim Fg As Double, Fs As Double
    Dim Rho As Double, R As Double, g As Double
    Dim MyDmy As Double
    
    If Wilcock Then
        MessageOnWelcome "Calculating with Wilcock (2001) two-fraction model." & _
            vbLf & vbLf & "Please wait..."
            
        Set InSt = Worksheets("Input")
        
        Nsp = InSt.Cells(7, 1).Value
        Rough = 2 * InSt.Cells(13, 2).Value / 1000 ' twice of D65, convert from mm to m
        Slope = InSt.Cells(5, 2).Value
        Fg = InSt.Cells(12, 1).Value
        Fs = 1 - Fg
        Rho = 1000
        R = 1.65
        g = 9.81
        
        Canceled = False
        GoGetTaurGTaurS Nsp, Rough, Fg, R, g, Slope, TaurG, TaurS
        If Canceled Then
            EndMessageOnWelcome 1
            GoTo 10
        End If
        If InSt.Cells(6, 1).Value = "(A)" Then 'single discharge
            Ncal = 1
        Else
            Ncal = 26
        End If
        If InSt.Cells(1, 1).Value = "XS" Then 'cross section is used
            For i = 1 To Ncal
                Qw = InSt.Cells(i, 16).Value
                GetDepthWithDischargeManningStrickler Qw, Slope, Rough, InSt.Cells(3, 1).Value, _
                    InSt.Cells(3, 2).Value, Qwc, Qw1, Qw2, H(i), Rh, Area, g
                Ustar = (g * Rh * Slope) ^ 0.5
                tau = Rho * Ustar ^ 2
                If TaurG > 0 Then
                    If tau > TaurG Then
                        QG(i) = Fg * Area * Slope * Ustar / R * 11.2 * (1 - 0.846 * TaurG / tau) ^ 4.5
                    Else
                        QG(i) = Fg * Area * Slope * Ustar / R * 0.0025 * (tau / TaurG) ^ 14.2
                    End If
                End If
                If TaurS > 0 Then
                    MyDmy = 1 - 0.846 * (TaurS / tau) ^ 0.5
                    If MyDmy > 0 Then
                        Qs(i) = Fs * Area * Slope * Ustar / R * 11.2 * MyDmy ^ 4.5
                    Else
                        Qs(i) = 0
                    End If
                End If
            Next 'i=1 to Ncal
        Else 'bankfull width is used
            Width = InSt.Cells(1, 2).Value
            For i = 1 To Ncal
                Qw = InSt.Cells(i, 16).Value
                H(i) = (Qw / 8.1 / Width) ^ 0.6 * Rough ^ 0.1 / (g * Slope) ^ 0.3
                Ustar = (g * H(i) * Slope) ^ 0.5
                tau = Rho * Ustar ^ 2
                If TaurG > 0 Then
                    If tau > TaurG Then
                        QG(i) = Fg * Width * Ustar ^ 3 / R / g * 11.2 * (1 - 0.846 * TaurG / tau) ^ 4.5
                    Else
                        QG(i) = Fg * Width * Ustar ^ 3 / R / g * 0.0025 * (tau / TaurG) ^ 14.2
                    End If
                End If
                If TaurS > 0 Then _
                    MyDmy = 1 - 0.846 * (TaurS / tau) ^ 0.5
                    If MyDmy > 0 Then
                        Qs(i) = Fs * Width * Ustar ^ 3 / R / g * 11.2 * MyDmy ^ 4.5
                    Else
                        Qs(i) = 0
                    End If
            Next 'i=1 to Ncal
        End If
        EndMessageOnWelcome 1
        PresentResultsForWilcock01 QG, Qs, H, TaurG, TaurS
        ModifyMenu
    End If
10  WilcockAndCroweSurfaceBased 1
End Sub

Sub WilcockAndCroweSurfaceBased(dmy As Integer)
    Dim InSt As Worksheet, StSt As Worksheet
    Dim Qs(27) As Double, Phi(27) As Double
    Dim Qw As Double, Slope As Double
    Dim fRange As Range, dRange As Range
    Dim Dsg As Double, STD As Double, H(26) As Double, D65 As Double, Fs As Double
    Dim Psi(21) As Double, f(21) As Double, p(21) As Double, Ap(21) As Double
    Dim Taursg As Double
    Dim Nsize As Long 'number of grain size fractions
    Dim Ncal As Long ' number of calculations
    Dim Rho As Double, R As Double, g As Double, Rough As Double
    Dim i As Long, j As Long, k As Long
    Dim Sum As Double, Dura As Double
       
    If Wilcock03 Then
        MessageOnWelcome "Calculating with the surface-based bedload equation of " & _
            "Wilcock and Crowe (2003)." & vbLf & vbLf & "Please wait..."
            
        Rho = 1000: R = 1.65: g = 9.81
        
        Set InSt = Worksheets("Input")
        Set StSt = Worksheets("Storage")
        
        Nsize = 0
        Do While Not IsEmpty(InSt.Cells(Nsize + 1, 5))
            Nsize = Nsize + 1
        Loop
        Nsize = Nsize - 1
        
        Set dRange = Range(StSt.Cells(1, 3), StSt.Cells(Nsize + 1, 3))
        Set fRange = Range(StSt.Cells(1, 4), StSt.Cells(Nsize + 1, 4))
        
        For i = 1 To Nsize + 1
            Psi(i) = Log(InSt.Cells(i, 5).Value) / Log(2)
        Next
        For i = 1 To Nsize
            Ap(i) = 0
            f(i) = Abs(InSt.Cells(i + 1, 6).Value - InSt.Cells(i, 6).Value) / 100
        Next
        
        GetGeometricMeanGrainSizeAndArithmeticStandardDeviation Nsize, Psi, f, Dsg, STD
        
        D65 = InSt.Cells(13, 2).Value / 1000 ' mm ==> m
        Rough = 2 * D65
        Fs = 1 - InSt.Cells(12, 1).Value
        Slope = InSt.Cells(5, 2).Value
        
        Taursg = WilcockTauRsg(Rho, R, g, Dsg, Fs)
            
        If InSt.Cells(6, 1).Value = "(A)" Then 'single discharge
            Ncal = 1
        Else ' multiple (26) discharges
            Ncal = 26
        End If
        If InSt.Cells(1, 1).Value = "XS" Then 'cross section is used
            For i = 1 To Ncal
                Qw = InSt.Cells(i, 16).Value
                CrossSectionWilcock03TransportRate Nsize, Rho, R, g, Dsg, Qw, Taursg, Rough, Slope, _
                    Psi, f, p, Qs(i), Phi(i), H(i)
                For j = 1 To Nsize
                    StSt.Cells(i, j + 26).Value = Qs(i) * p(j)
                Next
                If Left(InSt.Cells(6, 1).Value, 2) = "(C" Then
                    If i = 1 Then
                        Dura = 0.5 * (InSt.Cells(1, 17).Value - InSt.Cells(2, 17).Value) / 100
                    ElseIf i = 26 Then
                        Dura = 0.5 * (InSt.Cells(25, 17).Value - InSt.Cells(26, 17).Value) / 100
                    Else
                        Dura = 0.5 * (InSt.Cells(i - 1, 17).Value - InSt.Cells(i + 1, 17).Value) / 100
                    End If
                    For j = 1 To Nsize
                        Ap(j) = Ap(j) + Qs(i) * p(j) * Dura
                    Next
                End If
            Next 'i=1 to Ncal
        Else ' bankfull width is used
            For i = 1 To Ncal
                Qw = InSt.Cells(i, 16).Value
                BasicWilcock03TransportRate Nsize, Rho, R, g, Dsg, Qw, Taursg, Rough, Slope, _
                    Psi, f, p, Qs(i), Phi(i), H(i)
                For j = 1 To Nsize
                    StSt.Cells(i, j + 26).Value = Qs(i) * p(j)
                Next
                If Left(InSt.Cells(6, 1).Value, 2) = "(C" Then
                    If i = 1 Then
                        Dura = 0.5 * (InSt.Cells(1, 17).Value - InSt.Cells(2, 17).Value) / 100
                    ElseIf i = 26 Then
                        Dura = 0.5 * (InSt.Cells(25, 17).Value - InSt.Cells(26, 17).Value) / 100
                    Else
                        Dura = 0.5 * (InSt.Cells(i - 1, 17).Value - InSt.Cells(i + 1, 17).Value) / 100
                    End If
                    For j = 1 To Nsize
                        Ap(j) = Ap(j) + Qs(i) * p(j) * Dura
                    Next
                End If
            Next 'i=1 to Ncal
        End If
        If Left(InSt.Cells(6, 1).Value, 2) = "(C" Then
            Sum = 0
            For j = 1 To Nsize
                Sum = Sum + Ap(j)
            Next
            If Sum > 0 Then
                For j = 1 To Nsize
                    Ap(j) = Ap(j) / Sum
                Next
            End If
        End If
        If InSt.Cells(6, 1).Value = "(A)" Then
            For j = 1 To Nsize
                Ap(j) = p(j)
            Next
        End If
        EndMessageOnWelcome 1
        PresentResultsForWilcock03Results Qs, Phi, H, p
        ModifyMenu
    End If
    BakkeCalibratedParker82 1
End Sub

Sub BakkeCalibratedParker82(dmy As Integer)
    Dim Ncal As Long, i As Long, j As Long, k As Long
    Dim InSt As Worksheet, StSt As Worksheet
    Dim Qs(27) As Double, H(27) As Double, Phi50(27) As Double
    Dim TauR50Star As Double, Exponent As Double
    Dim R As Double, g As Double, Dk As Double, D50mm As Double, Width As Double
    Dim f(21) As Double, p(21) As Double, Psi(21) As Double, Ap(21) As Double
    Dim Slope As Double, Dura As Double, Sum As Double, Qw As Double
    Dim Nsize As Long, Nsp As Long
    Dim SizeRange As Range, CalculatedPctRange As Range, PctRange As Range, D50 As Double
    
    If Bakke Then
        i = MsgBox("Preparing to apply Bakke (1999) ..." & vbLf & vbLf & _
            "The original procedure of Bakke (1999) applies the bedload equation of " & _
            "Parker and Klingeman (1982).  This software offers you the choice of " & _
            "Parker and Klingeman (1982) or the slightly different Parker-Klingeman-McLean (1982) " & _
            "as an alternative." & vbLf & vbLf & _
            "Use the altrnative? (Click Yes to use Parker-Klingeman-McLean and No to use " & _
            "Parker-Klingeman)", vbYesNo + vbQuestion, "Select a equation")
        If i = vbYes Then
            PKorPKM = "PKM"
        ElseIf i = vbNo Then
            PKorPKM = "PK"
        Else
            Exit Sub
        End If
        Worksheets("Input").Cells(14, 1).Value = PKorPKM
        MessageOnWelcome "Calculating with Bakke et al. (1999)" & _
            vbLf & vbLf & "This may take a while!  Please wait..."
        
        Set InSt = Worksheets("Input"): Set StSt = Worksheets("Storage")
        
        Nsize = 0
        Do While Not IsEmpty(InSt.Cells(Nsize + 2, 7))
            Nsize = Nsize + 1
        Loop
        
        Set SizeRange = Range(InSt.Cells(1, 7), InSt.Cells(Nsize + 2, 7))
        Set CalculatedPctRange = Range(StSt.Cells(1, 7), StSt.Cells(Nsize + 2, 7))
                
        If IsEmpty(InSt.Cells(1, 2)) Then
            Width = InSt.Cells(4, 2).Value - InSt.Cells(4, 1).Value
        Else
            Width = InSt.Cells(1, 2).Value
        End If
        
        For i = 1 To Nsize
            Ap(i) = 0
            f(i) = Abs(InSt.Cells(i + 1, 8).Value - InSt.Cells(i, 8).Value) / 100
        Next
        For i = 1 To Nsize + 1
            Psi(i) = Log(InSt.Cells(i, 7).Value) / Log(2)
        Next
        
        Nsp = InSt.Cells(7, 1).Value
        
        R = 1.65: g = 9.81: Dk = 10.7
        D50mm = InSt.Cells(13, 1).Value
        Slope = InSt.Cells(5, 2).Value
        
        Worksheets("Storage").Columns("E:E").ClearContents
        For i = 1 To Nsp
            If InSt.Cells(i, 19).Value = 0 Then
                Worksheets("Storage").Cells(i, 5).Value = 1E+20
            Else
                Set PctRange = Range(InSt.Cells(i, 21), InSt.Cells(i, 21 + Nsize))
                GetCharacteristicGrainSizeinMM PctRange, SizeRange, 50, Sum
                Worksheets("Storage").Cells(i, 5).Value = Sum
            End If
        Next
        
        GoGetTau50StarAndExponent TauR50Star, Exponent, R, g, Dk, D50mm, Width, f, p, Psi, _
            Slope, Nsize, Nsp

        If InSt.Cells(6, 1).Value = "(A)" Then 'single discharge
            Ncal = 1
        Else
            Ncal = 26
        End If
        If InSt.Cells(1, 1).Value = "XS" Then 'cross section is used
            For i = 1 To Ncal
                Qw = InSt.Cells(i, 16).Value
                k = ParkerklingemanEquationWithCrossSection(Exponent, TauR50Star, _
                    R, g, Dk, D50mm, Qw, f, p, Psi, Qs(i), H(i), Phi50(i), Slope, Nsize, "Bakke")
                For k = 1 To Nsize
                    StSt.Cells(i, k + 26).Value = Qs(i) * p(k)
                Next
                If Left(InSt.Cells(6, 1).Value, 2) = "(C" Then
                    If i = 1 Then
                        Dura = 0.5 * (InSt.Cells(1, 17).Value - InSt.Cells(2, 17).Value) / 100
                    ElseIf i = 26 Then
                        Dura = 0.5 * (InSt.Cells(25, 17).Value - InSt.Cells(26, 17).Value) / 100
                    Else
                        Dura = 0.5 * (InSt.Cells(i - 1, 17).Value - InSt.Cells(i + 1, 17).Value) / 100
                    End If
                    For j = 1 To Nsize
                        Ap(j) = Ap(j) + Qs(i) * p(j) * Dura
                    Next
                End If
                
                If InSt.Cells(1, 8).Value < 1 Then 'increasing size
                    D50 = 0
                    StSt.Cells(1, 7).Value = D50
                    For j = 1 To Nsize
                        D50 = D50 + p(j) * 100
                        StSt.Cells(j + 1, 7).Value = Format(D50, "##0.#")
                    Next
                Else 'decreasing size
                    D50 = 100
                    StSt.Cells(1, 7).Value = D50
                    For j = 1 To Nsize
                        D50 = D50 - p(j) * 100
                        StSt.Cells(j + 1, 7).Value = Format(D50, "##0.#")
                    Next
                End If
                GetCharacteristicGrainSizeinMM CalculatedPctRange, SizeRange, 50, D50
                StSt.Cells(i, 10).Value = Format(D50, "##0.#")
            Next 'i=1 to Ncal
        Else 'bankfull width is used
            For i = 1 To Ncal
                Qw = InSt.Cells(i, 16).Value
                k = ParkerKlingemanBasicEquation(Exponent, TauR50Star, R, g, Dk, D50mm, Width, _
                    Qw, f, p, Psi, Qs(i), H(i), Phi50(i), Slope, Nsize, "Bakke")
                For k = 1 To Nsize
                    StSt.Cells(i, k + 26).Value = Qs(i) * p(k)
                Next
                If Left(InSt.Cells(6, 1).Value, 2) = "(C" Then
                    If i = 1 Then
                        Dura = 0.5 * (InSt.Cells(1, 17).Value - InSt.Cells(2, 17).Value) / 100
                    ElseIf i = 26 Then
                        Dura = 0.5 * (InSt.Cells(25, 17).Value - InSt.Cells(26, 17).Value) / 100
                    Else
                        Dura = 0.5 * (InSt.Cells(i - 1, 17).Value - InSt.Cells(i + 1, 17).Value) / 100
                    End If
                    For j = 1 To Nsize
                        Ap(j) = Ap(j) + Qs(i) * p(j) * Dura
                    Next
                End If
                
                If InSt.Cells(1, 8).Value < 1 Then 'increasing size
                    D50 = 0
                    StSt.Cells(1, 7).Value = D50
                    For j = 1 To Nsize
                        D50 = D50 + p(j) * 100
                        StSt.Cells(j + 1, 7).Value = Format(D50, "##0.#")
                    Next
                Else 'decreasing size
                    D50 = 100
                    StSt.Cells(1, 7).Value = D50
                    For j = 1 To Nsize
                        D50 = D50 - p(j) * 100
                        StSt.Cells(j + 1, 7).Value = Format(D50, "##0.#")
                    Next
                End If
                GetCharacteristicGrainSizeinMM CalculatedPctRange, SizeRange, 50, D50
                StSt.Cells(i, 10).Value = Format(D50, "##0.#")
            Next 'i=1 to Ncal
        End If
        If Left(InSt.Cells(6, 1).Value, 2) = "(C" Then
            Sum = 0
            For j = 1 To Nsize
                Sum = Sum + Ap(j)
            Next
            If Sum > 0 Then
                For j = 1 To Nsize
                    Ap(j) = Ap(j) / Sum
                Next
            End If
        End If
        If InSt.Cells(6, 1).Value = "(A)" Then
            For j = 1 To Nsize
                Ap(j) = p(j)
            Next
        End If
        EndMessageOnWelcome 1
        PresentResultsForParkerKlingeman82 Qs, Phi50, H, Ap, 1
        ModifyMenu
    End If
End Sub

Sub GetDepthWithDischargeManningsn(Qw As Double, Slope As Double, nc As Double, _
    nl As Double, nr As Double, Qwc As Double, Qwl As Double, Qwr As Double, _
    H As Double, Rh As Double, Area As Double)

    Dim Err As Double 'relative error in iteration
    Dim Hup As Double, Hlw As Double ' upper and lower limit of water depth, with bisect method
    Dim Rhl As Double, Areal As Double
    Dim Rhr As Double, Arear As Double
    Dim dmy As Double
    
    Hup = Worksheets("Input").Cells(51, 9).Value * 3 'three times of the maximum depth provided in cross section
    Hlw = 0
    H = Hup
    GetRhAndAreaFromDepth H, Rh, Area, Rhl, Areal, Rhr, Arear
    Err = 1
    Do While Err > 0.00001
        If nl > 0.0001 Or nl < 10 Then
            Qwl = Areal * Rhl ^ (2 / 3) * Slope ^ 0.5 / nl
        Else
            Qwl = 0
        End If
        If nr > 0.0001 Or nr < 10 Then
            Qwr = Arear * Rhr ^ (2 / 3) * Slope ^ 0.5 / nr
        Else
            Qwr = 0
        End If
        Qwc = Area * Rh ^ (2 / 3) * Slope ^ 0.5 / nc
        Err = Abs((Qwl + Qwr + Qwc - Qw) / Qw)
        If (Qwl + Qwr + Qwc) > Qw Then ' decrease H
            Hup = H
            H = 0.5 * (Hup + Hlw)
        ElseIf (Qwl + Qwr + Qwc) < Qw Then ' increase H
            Hlw = H
            H = 0.5 * (Hup + Hlw)
        End If
        GetRhAndAreaFromDepth H, Rh, Area, Rhl, Areal, Rhr, Arear
    Loop

End Sub


