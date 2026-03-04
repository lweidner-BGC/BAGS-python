Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

Sub GoGetTaurGTaurS(Nsp As Long, Rough As Double, Fg As Double, R As Double, g As Double, _
    Slope As Double, TaurG As Double, TaurS As Double)

    Dim Fs As Double, Width As Double
    Dim InSt As Worksheet
    Dim i As Long, j As Long
    Dim FGP As Double, FSP As Double ' half of derivative of square of errors
    Dim Qw As Double, QGs As Double, QGc As Double, QSs As Double, QSc As Double
    Dim lnQGp As Double, lnQSp As Double 'derivatives of ln(QGc) and ln(QSc)
    Dim Rho As Double, Ustar As Double, H As Double, tau As Double
    Dim QGError As Double, QSError As Double 'relative errors in iteration
    Dim Tolerance As Double 'tolerance in relative errors
    Dim Mn1 As Double, Mn2 As Double, Qw1 As Double, Qw2 As Double, Qwc As Double
    Dim Rh As Double, Area As Double
    Dim kount As Long

    Set InSt = Worksheets("Input")
    
    Rho = 1000 'density of water, in kg/m3
    
    TaurG = Rho * R * g * Rough * 0.04  'initial guess
    TaurS = Rho * R * g / 1000 * 0.1  ' initial guess, assuming 1 mm grain size
    
    InitializingProgressBar 1
    If InSt.Cells(1, 1).Value <> "XS" Then
        Width = InSt.Cells(1, 2).Value
        If Fg > 0 Then GoGetTaurG Nsp, Rough, Fg, R, g, Slope, Width, TaurG
        If Fg < 1 Then GoGetTaurS Nsp, Rough, Fg, R, g, Slope, Width, TaurS
    Else
        Width = InSt.Cells(4, 2).Value - InSt.Cells(4, 1).Value
        If Fg > 0 Then GoGetTaurG Nsp, Rough, Fg, R, g, Slope, Width, TaurG
        If Fg < 1 Then GoGetTaurS Nsp, Rough, Fg, R, g, Slope, Width, TaurS
        
        Width = -1
        If Fg > 0 Then GoGetTaurG Nsp, Rough, Fg, R, g, Slope, Width, TaurG
        If Fg < 1 Then GoGetTaurS Nsp, Rough, Fg, R, g, Slope, Width, TaurS
    End If

End Sub

Private Sub GoGetTaurG(Nsp As Long, Rough As Double, Fg As Double, R As Double, g As Double, _
    Slope As Double, Width As Double, TaurG As Double)

    If Not FirstGoGetTaurG(Nsp, Rough, Fg, R, g, Slope, Width, TaurG) Then _
        SecondGoGetTaurG Nsp, Rough, Fg, R, g, Slope, Width, TaurG
End Sub

Private Sub GoGetTaurS(Nsp As Long, Rough As Double, Fg As Double, R As Double, g As Double, _
    Slope As Double, Width As Double, TaurG As Double)

    If Not FirstGoGetTaurS(Nsp, Rough, Fg, R, g, Slope, Width, TaurG) Then _
        SecondGoGetTaurS Nsp, Rough, Fg, R, g, Slope, Width, TaurG
End Sub

Private Sub SecondGoGetTaurG(Nsp As Long, Rough As Double, Fg As Double, R As Double, g As Double, _
    Slope As Double, Width As Double, TaurG As Double)
    
    Dim InSt As Worksheet, MySt As Worksheet
    Dim i As Long, j As Long
    Dim lnQGs As Double, lnQGc As Double
    Dim Rho As Double, Ustar As Double, H As Double, Qw As Double, tau As Double
    Dim Mn1 As Double, Mn2 As Double, Qw1 As Double, Qw2 As Double, Qwc As Double
    Dim Rh As Double, Area As Double
    Dim QGError As Double, Tolerance As Double
    Dim uTaurG As Double, lTaurG As Double, cTaurG
    Dim uT As Double, lT As Double 'absolute upper and lower limit for TaurG
    Dim dmy As Double, lnC As Double
    Dim SumSqt As Double
    Dim ErrOut As Label
    Dim TempStor1(1 To 51) As Double, TempStor2(1 To 51) As Double
    
    On Error GoTo ErrOut
    
    lnC = Log(2650) + Log(60)
    Tolerance = 0.001
    Rho = 1000 'density of water, in kg/m3
    Set InSt = Worksheets("Input")
    Set MySt = Worksheets("Storage")
        
    If Width < 0 Then ' cross section is used
        Mn1 = InSt.Cells(3, 1).Value
        Mn2 = InSt.Cells(3, 2).Value
    End If
    
    QGError = 0: cTaurG = TaurG: uTaurG = 10 * TaurG: lTaurG = TaurG / 10: TaurG = lTaurG
    lT = -1: uT = -1
    For i = 1 To Nsp
        If Not InSt.Cells(i, 40).Value Then _
            QGError = QGError + InSt.Cells(i, 19).Value * InSt.Cells(i, 20).Value
    Next

    If QGError = 0 Then ' no gravel transport samples are supplied
        TaurG = 0
        Exit Sub
    Else
        QGError = 1  ' allowing for TaurG calculation
    End If

    Do While QGError > Tolerance
        For j = 1 To 51
            UpdatingProgressBar 1
            If j = 1 Then
                TaurG = lTaurG
            Else
                TaurG = TaurG * (uTaurG / lTaurG) ^ (1 / 50)
            End If
            SumSqt = 0
            For i = 1 To Nsp
                If InSt.Cells(i, 40).Value Then GoTo 10
                Qw = InSt.Cells(i, 18).Value
                If QGError > Tolerance Then
                    dmy = InSt.Cells(i, 19).Value * InSt.Cells(i, 20).Value
                    If dmy > 0 Then
                        lnQGs = Log(dmy)
                    Else
                        lnQGs = 1E+20
                    End If
                End If
                If Width < 0 Then 'using cross section
                    GetDepthWithDischargeManningStrickler _
                        Qw, Slope, Rough, Mn1, Mn2, Qwc, Qw1, Qw2, H, Rh, Area, g
                    Ustar = (g * Rh * Slope) ^ 0.5
                    tau = Rho * Ustar ^ 2
                    If QGError > Tolerance Then
                        If tau > TaurG Then
                            lnQGc = Log(Fg * Area * Slope * Ustar / R * 11.2) + _
                                4.5 * Log(1 - 0.846 * TaurG / tau) + lnC
                        Else
                            lnQGc = Log(Fg * Area * Slope * Ustar / R * 0.0025) + _
                                14.2 * Log(tau / TaurG) + lnC
                        End If
                    End If
                Else 'using bankfull width
                    H = (Qw / 8.1 / Width) ^ 0.6 * Rough ^ 0.1 / (g * Slope) ^ 0.3
                    Ustar = (g * H * Slope) ^ 0.5
                    tau = Rho * Ustar ^ 2
                    If QGError > Tolerance Then
                        If tau > TaurG Then
                            lnQGc = Log(Fg * Width * Ustar ^ 3 / R / g * 11.2) + _
                                4.5 * Log(1 - 0.846 * TaurG / tau) + lnC
                        Else
                            lnQGc = Log(Fg * Width * Ustar ^ 3 / R / g * 0.0025) + _
                                14.2 * Log(tau / TaurG) + lnC
                        End If
                    End If
                End If
                If lnQGs < 1E+19 Then
                    SumSqt = SumSqt + (lnQGc - lnQGs) ^ 2
                End If
10          Next 'i
            TempStor1(j) = TaurG
            TempStor2(j) = SumSqt
        Next 'j
        SumSqt = Application.WorksheetFunction.Min(TempStor2)
        If Abs(TempStor2(1) - SumSqt) < 0.00001 Then
            cTaurG = TempStor1(1)
            uT = TempStor1(2)
            uTaurG = (cTaurG * uT) ^ 0.5
            If lT < 0 Then
                lTaurG = cTaurG / 1.35
            Else
                lTaurG = (cTaurG * lT) ^ 0.5
            End If
        ElseIf Abs(TempStor2(51) - SumSqt) < 0.00001 Then
            cTaurG = TempStor1(51)
            lT = TempStor1(50)
            lTaurG = (cTaurG * lT) ^ 0.5
            If uT < 0 Then
                uTaurG = cTaurG * 1.5
            Else
                uTaurG = (cTaurG * uT) ^ 0.5
            End If
        Else
            For j = 1 To 51
                If Abs(TempStor2(j) - SumSqt) < 0.00001 Then Exit For
            Next
            cTaurG = TempStor1(j)
            lTaurG = TempStor1(j - 1)
            uTaurG = TempStor1(j + 1)
            uT = uTaurG
            lT = lTaurG
        End If
        TaurG = (uTaurG * lTaurG) ^ 0.5
        QGError = (uTaurG - lTaurG) / TaurG
    Loop
    Exit Sub
ErrOut:
    MsgBox "Calculation of reference shear stress for gravel does not converge!"
End Sub

Private Sub SecondGoGetTaurS(Nsp As Long, Rough As Double, Fg As Double, R As Double, g As Double, _
    Slope As Double, Width As Double, TaurS As Double)
    
    Dim Fs As Double
    Dim InSt As Worksheet, MySt As Worksheet
    Dim i As Long, j As Long
    Dim lnQSs As Double, lnQSc As Double
    Dim Rho As Double, Ustar As Double, H As Double, Qw As Double, tau As Double
    Dim Mn1 As Double, Mn2 As Double, Qw1 As Double, Qw2 As Double, Qwc As Double
    Dim Rh As Double, Area As Double
    Dim QSError As Double, Tolerance As Double
    Dim uTaurS As Double, lTaurS As Double, cTaurS
    Dim uT As Double, lT As Double 'absolute upper and lower limit for TaurS
    Dim dmy As Double, lnC As Double
    Dim SumSqt As Double
    Dim ErrOut As Label
    Dim TempStor1(1 To 51) As Double, TempStor2(1 To 51) As Double
    
    On Error GoTo ErrOut
    
    Fs = 1 - Fg
    lnC = Log(2650) + Log(60)
    Tolerance = 0.001
    Rho = 1000 'density of water, in kg/m3
    Set InSt = Worksheets("Input")
    Set MySt = Worksheets("Storage")
    
    If Width < 0 Then ' cross section is used
        Mn1 = InSt.Cells(3, 1).Value
        Mn2 = InSt.Cells(3, 2).Value
    End If
    
    QSError = 0: cTaurS = TaurS: uTaurS = 10 * TaurS: lTaurS = TaurS / 10: TaurS = lTaurS
    lT = -1: uT = -1
    For i = 1 To Nsp
        If Not InSt.Cells(i, 40).Value Then _
            QSError = QSError + InSt.Cells(i, 19).Value * (1 - InSt.Cells(i, 20).Value)
    Next

    If QSError = 0 Then ' no sand transport sampling provided
        TaurS = 0
        Exit Sub
    Else
        QSError = 1  ' allowing for TaurG calculation
    End If

    Do While QSError > Tolerance
        For j = 1 To 51
            UpdatingProgressBar 1
            If j = 1 Then
                TaurS = lTaurS
            Else
                TaurS = TaurS * (uTaurS / lTaurS) ^ (1 / 50)
            End If
            SumSqt = 0
            For i = 1 To Nsp
                If InSt.Cells(i, 40).Value Then GoTo 10
                Qw = InSt.Cells(i, 18).Value
                If QSError > Tolerance Then
                    dmy = InSt.Cells(i, 19).Value * (1 - InSt.Cells(i, 20).Value)
                    If dmy > 0 Then
                        lnQSs = Log(dmy)
                    Else
                        lnQSs = 1E+20
                    End If
                End If
                If Width < 0 Then 'using cross section
                    GetDepthWithDischargeManningStrickler _
                        Qw, Slope, Rough, Mn1, Mn2, Qwc, Qw1, Qw2, H, Rh, Area, g
                    Ustar = (g * Rh * Slope) ^ 0.5
                    tau = Rho * Ustar ^ 2
                    If QSError > Tolerance Then
                        dmy = 1 - 0.846 * (TaurS / tau) ^ 0.5
                        If dmy > 0 Then
                            lnQSc = Log(Fs * Area * Slope * Ustar / R * 11.2) + 4.5 * Log(dmy) + lnC
                        Else
                            lnQSc = 1E+20
                        End If
                    End If
                Else 'using bankfull width
                    H = (Qw / 8.1 / Width) ^ 0.6 * Rough ^ 0.1 / (g * Slope) ^ 0.3
                    Ustar = (g * H * Slope) ^ 0.5
                    tau = Rho * Ustar ^ 2
                    If QSError > Tolerance Then
                        dmy = 1 - 0.846 * (TaurS / tau) ^ 0.5
                        If dmy > 0 Then
                            lnQSc = Log(Fs * Width * Ustar ^ 3 / R / g * 11.2) + 4.5 * Log(dmy) + lnC
                        Else
                            lnQSc = 1E+20
                        End If
                    End If
                End If
                If lnQSs < 1E+19 Then
                    SumSqt = SumSqt + (lnQSc - lnQSs) ^ 2
                End If
10          Next
            TempStor1(j) = TaurS
            TempStor2(j) = SumSqt
        Next
        SumSqt = Application.WorksheetFunction.Min(TempStor2)
        If Abs(TempStor2(1) - SumSqt) < 0.00001 Then
            cTaurS = TempStor1(1)
            uT = TempStor1(2)
            uTaurS = (cTaurS * uT) ^ 0.5
            If lT < 0 Then
                lTaurS = cTaurS / 1.35
            Else
                lTaurS = (cTaurS * lT) ^ 0.5
            End If
        ElseIf Abs(TempStor2(51) - SumSqt) < 0.00001 Then
            cTaurS = TempStor1(51)
            lT = TempStor1(50)
            lTaurS = (cTaurS * lT) ^ 0.5
            If uT < 0 Then
                uTaurS = cTaurS * 1.5
            Else
                uTaurS = (cTaurS * uT) ^ 0.5
            End If
        Else
            For j = 1 To 51
                If Abs(TempStor2(j) - SumSqt) < 0.00001 Then Exit For
            Next
            cTaurS = TempStor1(j)
            lTaurS = TempStor1(j - 1)
            uTaurS = TempStor1(j + 1)
            lT = lTaurS
            uT = uTaurS
        End If
        TaurS = (uTaurS * lTaurS) ^ 0.5
        QSError = (uTaurS - lTaurS) / TaurS
    Loop
    Exit Sub
ErrOut:
    MsgBox "Calculation of reference shear stress for sand does not converge!"
End Sub

Private Function FirstGoGetTaurG(Nsp As Long, Rough As Double, Fg As Double, R As Double, g As Double, _
    Slope As Double, Width As Double, TaurG As Double) As Boolean
    Dim InSt As Worksheet
    Dim i As Long, j As Long, kount As Long
    Dim lnQGs As Double, lnQGc As Double
    Dim lnQGp As Double, lnQGpp As Double
    Dim FGP As Double, FGPp As Double
    Dim Rho As Double, Ustar As Double, H As Double, Qw As Double, tau As Double
    Dim QGError As Double, Tolerance As Double
    Dim Mn1 As Double, Mn2 As Double, Qw1 As Double, Qw2 As Double, Qwc As Double
    Dim Rh As Double, Area As Double
    Dim uTaurG As Double, lTaurG As Double
    Dim dmy As Double, lnC As Double
    Dim SumSqt As Double
    Dim ErrOut As Label
    
    On Error GoTo ErrOut
    
    lnC = Log(2650) + Log(60)
    Tolerance = 0.00001
    Rho = 1000 'density of water, in kg/m3
    Set InSt = Worksheets("Input")
    
    If Width < 0 Then ' cross section is used
        Mn1 = InSt.Cells(3, 1).Value
        Mn2 = InSt.Cells(3, 2).Value
    End If
    
    QGError = 0: uTaurG = -1: lTaurG = -1
    For i = 1 To Nsp
        If Not InSt.Cells(i, 40).Value Then _
            QGError = QGError + InSt.Cells(i, 19).Value * InSt.Cells(i, 20).Value
    Next

    If QGError = 0 Then
        TaurG = 0
        Exit Function
    Else
        QGError = 1  ' allowing for TaurG calculation
    End If

    kount = 0
    Do While QGError > Tolerance
        kount = kount + 1
        UpdatingProgressBar 1
        FGP = 0: FGPp = 0
        SumSqt = 0
        For i = 1 To Nsp
            If InSt.Cells(i, 40).Value Then GoTo 10
            Qw = InSt.Cells(i, 18).Value
            If QGError > Tolerance Then
                dmy = InSt.Cells(i, 19).Value * InSt.Cells(i, 20).Value
                If dmy > 0 Then
                    lnQGs = Log(dmy)
                Else
                    lnQGs = 1E+20
                End If
            End If
            If Width < 0 Then 'using cross section
                GetDepthWithDischargeManningStrickler _
                    Qw, Slope, Rough, Mn1, Mn2, Qwc, Qw1, Qw2, H, Rh, Area, g
                Ustar = (g * Rh * Slope) ^ 0.5
                tau = Rho * Ustar ^ 2
                If QGError > Tolerance Then
                    If tau > TaurG Then
                        lnQGc = Log(Fg * Area * Slope * Ustar / R * 11.2) + _
                            4.5 * Log(1 - 0.846 * TaurG / tau) + lnC
                        lnQGp = -4.5 * 0.846 / (tau - 0.846 * TaurG)
                        lnQGpp = -4.5 * (0.846 / (tau - 0.846 * TaurG)) ^ 2
                    Else
                        lnQGc = Log(Fg * Area * Slope * Ustar / R * 0.0025) + _
                            14.2 * Log(tau / TaurG) + lnC
                        lnQGp = -14.2 / TaurG
                        lnQGpp = 14.2 / TaurG ^ 2
                    End If
                End If
            Else 'using bankfull width
                H = (Qw / 8.1 / Width) ^ 0.6 * Rough ^ 0.1 / (g * Slope) ^ 0.3
                Ustar = (g * H * Slope) ^ 0.5
                tau = Rho * Ustar ^ 2
                If QGError > Tolerance Then
                    If tau > TaurG Then
                        lnQGc = Log(Fg * Width * Ustar ^ 3 / R / g * 11.2) + _
                            4.5 * Log(1 - 0.846 * TaurG / tau) + lnC
                        lnQGp = -4.5 * 0.846 / (tau - 0.846 * TaurG)
                        lnQGpp = -4.5 * (0.846 / (tau - 0.846 * TaurG)) ^ 2
                    Else
                        lnQGc = Log(Fg * Width * Ustar ^ 3 / R / g * 0.0025) + _
                            14.2 * Log(tau / TaurG) + lnC
                        lnQGp = -14.2 / TaurG
                        lnQGpp = 14.2 / TaurG ^ 2
                    End If
                End If
            End If
            If lnQGs < 1E+19 And lnQGc < 1E+19 Then
                FGP = FGP + (lnQGc - lnQGs) * lnQGp
                FGPp = FGPp + lnQGp ^ 2 + (lnQGc - lnQGs) * lnQGpp
                SumSqt = SumSqt + (lnQGc - lnQGs) ^ 2
            End If
10      Next
        If FGP * FGPp < 0 Then ' need to increase TaurG
            lTaurG = TaurG
            If uTaurG < 0 Then
                TaurG = 1.05 * TaurG
            Else
                TaurG = 0.5 * (uTaurG + lTaurG)
                QGError = (uTaurG - lTaurG) / TaurG
            End If
        Else ' decrease TaurG
            uTaurG = TaurG
            If lTaurG < 0 Then
                TaurG = 0.95 * TaurG
            Else
                TaurG = 0.5 * (uTaurG + lTaurG)
                QGError = (uTaurG - lTaurG) / TaurG
            End If
        End If
        If kount = 200 And QGError > Tolerance Then GoTo ErrOut
    Loop
    FirstGoGetTaurG = True
    Exit Function
ErrOut:
    FirstGoGetTaurG = False
End Function

Private Function FirstGoGetTaurS(Nsp As Long, Rough As Double, Fg As Double, R As Double, g As Double, _
    Slope As Double, Width As Double, TaurS As Double) As Boolean
    Dim Fs As Double
    Dim InSt As Worksheet
    Dim i As Long, kount As Long
    Dim lnQSs As Double, lnQSc As Double
    Dim lnQSp As Double, lnQSpp As Double
    Dim FSP As Double, FSPp As Double
    Dim Rho As Double, Ustar As Double, H As Double, Qw As Double, tau As Double
    Dim QSError As Double, Tolerance As Double
    Dim Mn1 As Double, Mn2 As Double, Qw1 As Double, Qw2 As Double, Qwc As Double
    Dim Rh As Double, Area As Double
    Dim uTaurS As Double, lTaurS As Double
    Dim dmy As Double, lnC As Double
    Dim ErrOut As Label
    
    On Error GoTo ErrOut
    
    Fs = 1 - Fg
    lnC = Log(2650) + Log(60)
    Tolerance = 0.00001
    Rho = 1000 'density of water, in kg/m3
    Set InSt = Worksheets("Input")
    
    If Width < 0 Then ' cross section is used
        Mn1 = InSt.Cells(3, 1).Value
        Mn2 = InSt.Cells(3, 2).Value
    End If
    
    QSError = 0: uTaurS = -1: lTaurS = -1
    For i = 1 To Nsp
        If Not InSt.Cells(i, 40).Value Then _
            QSError = QSError + InSt.Cells(i, 19).Value * (1 - InSt.Cells(i, 20).Value)
    Next

    If QSError = 0 Then
        TaurS = 0
        Exit Function
    Else
        QSError = 1  ' allowing for TaurG calculation
    End If

    kount = 0
    Do While QSError > Tolerance
        kount = kount + 1
        UpdatingProgressBar 1
        FSP = 0: FSPp = 0
        For i = 1 To Nsp
            If InSt.Cells(i, 40).Value Then GoTo 10
            Qw = InSt.Cells(i, 18).Value
            If QSError > Tolerance Then
                dmy = InSt.Cells(i, 19).Value * (1 - InSt.Cells(i, 20).Value)
                If dmy > 0 Then
                    lnQSs = Log(dmy)
                Else
                    lnQSs = 1E+20
                End If
            End If
            If Width < 0 Then 'using cross section
                GetDepthWithDischargeManningStrickler _
                    Qw, Slope, Rough, Mn1, Mn2, Qwc, Qw1, Qw2, H, Rh, Area, g
                Ustar = (g * Rh * Slope) ^ 0.5
                tau = Rho * Ustar ^ 2
                If QSError > Tolerance Then
                    dmy = 1 - 0.846 * (TaurS / tau) ^ 0.5
                    If dmy > 0 Then
                        lnQSc = Log(Fs * Area * Slope * Ustar / R * 11.2) + 4.5 * Log(dmy) + lnC
                        lnQSp = -4.5 * 0.423 / ((tau * TaurS) ^ 0.5 - 0.846 * TaurS)
                        lnQSpp = 4.5 * 0.423 * (0.846 - 0.5 * (tau / TaurS) ^ 0.5) / _
                            ((tau * TaurS) ^ 0.5 - 0.846 * TaurS) ^ 2
                    Else
                        lnQSc = 1E+20
                        lnQSp = 0
                        lnQSpp = 0
                    End If
                End If
            Else 'using bankfull width
                H = (Qw / 8.1 / Width) ^ 0.6 * Rough ^ 0.1 / (g * Slope) ^ 0.3
                Ustar = (g * H * Slope) ^ 0.5
                tau = Rho * Ustar ^ 2
                If QSError > Tolerance Then
                    dmy = 1 - 0.846 * (TaurS / tau) ^ 0.5
                    If dmy > 0 Then
                        lnQSc = Log(Fs * Width * Ustar ^ 3 / R / g * 11.2) + 4.5 * Log(dmy) + lnC
                        lnQSp = -4.5 * 0.423 / ((tau * TaurS) ^ 0.5 - 0.846 * TaurS)
                        lnQSpp = 4.5 * 0.423 * (0.846 - 0.5 * (tau / TaurS) ^ 0.5) / _
                            ((tau * TaurS) ^ 0.5 - 0.846 * TaurS) ^ 2
                    Else
                        lnQSc = 1E+20
                        lnQSp = 0
                        lnQSpp = 0
                    End If
                End If
            End If
            If lnQSs < 1E+19 And lnQSc < 1E+19 Then
                FSP = FSP + (lnQSc - lnQSs) * lnQSp
                FSPp = FSPp + lnQSp ^ 2 + (lnQSc - lnQSs) * lnQSpp
            End If
10      Next
        If FSP * FSPp < 0 Then ' need to increase TaurS
            lTaurS = TaurS
            If uTaurS < 0 Then
                TaurS = 1.05 * TaurS
            Else
                TaurS = 0.5 * (uTaurS + lTaurS)
            End If
            QSError = (TaurS - lTaurS) / TaurS
        Else ' need to decrease TaurS
            uTaurS = TaurS
            If lTaurS < 0 Then
                TaurS = 0.95 * TaurS
            Else
                TaurS = 0.5 * (uTaurS + lTaurS)
            End If
            QSError = (uTaurS - TaurS) / TaurS
        End If
        If kount = 200 And QSError > Tolerance Then GoTo ErrOut
    Loop
    FirstGoGetTaurS = True
    Exit Function
ErrOut:
    FirstGoGetTaurS = False
End Function

Sub GetDepthWithDischargeManningStrickler _
    (Qw As Double, Slope As Double, Rough As Double, Mn1 As Double, Mn2 As Double, _
    Qwc As Double, Qw1 As Double, Qw2 As Double, H As Double, Rh As Double, _
    Area As Double, g As Double)
    
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
        If Mn1 > 0.0001 And Mn1 < 10 Then
            Qw1 = Area1 * Rh1 ^ (2 / 3) * Slope ^ 0.5 / Mn1
        Else
            Qw1 = 0
        End If
        If Mn2 > 0.0001 And Mn2 < 10 Then
            Qw2 = Area2 * Rh2 ^ (2 / 3) * Slope ^ 0.5 / Mn2
        Else
            Qw2 = 0
        End If
        Qwc = 8.1 * Area * (g * Rh * Slope) ^ 0.5 * (Rh / Rough) ^ (1 / 6)
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


Sub PresentResultsForWilcock01(QG() As Double, Qs() As Double, H() As Double, _
    TaurG As Double, TaurS As Double)
    Dim MyBk As Workbook, InSt As Worksheet
    Dim i As Long, Nsample As Long, nXS As Long
    Dim Rh As Double, Area As Double, Rh1 As Double, Area1 As Double, Rh2 As Double, Area2 As Double
    Dim xRange As Range, yRange As Range
    Dim cc As String
    
    Set InSt = Worksheets("Input")
    
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
        "Bedload transport equation used: The two-fraction equation of Wilcock (2001)"
    MyBk.Sheets(1).Cells(6, 2).Value = _
        "Input data are stored in worksheet ""Input"" and results are stored in worksheet ""Output""."
    MyBk.Sheets(1).Cells(8, 2).Value = _
        "Calculation was performed by " & Application.UserName & " on " & Date & "."

    If TaurG = 0 Or TaurS = 0 Then
        MyBk.Sheets(1).Cells(10, 2).Value = _
            "Gravel and/or sand transport rate(s) are not calculated due to:"
        MyBk.Sheets(1).Cells(11, 2).Value = _
            "    No bedload samples are provided for gravel and/or sand."
            
    End If

    MyBk.Sheets(2).Columns("A:A").ColumnWidth = 2
    MyBk.Sheets(2).Columns("B:B").ColumnWidth = 24
    MyBk.Sheets(2).Columns("C:C").ColumnWidth = 12
    MyBk.Sheets(2).Columns("D:D").ColumnWidth = 2
    MyBk.Sheets(2).Columns("E:E").ColumnWidth = 15
    MyBk.Sheets(2).Columns("F:F").ColumnWidth = 15
    MyBk.Sheets(2).Columns("G:G").ColumnWidth = 2
    MyBk.Sheets(2).Columns("H:H").ColumnWidth = 12
    MyBk.Sheets(2).Columns("I:I").ColumnWidth = 12
    MyBk.Sheets(2).Columns("J:J").ColumnWidth = 10
    MyBk.Sheets(2).Cells.Interior.ColorIndex = 2
    MyBk.Sheets(2).Cells.HorizontalAlignment = xlCenter
    MyBk.Sheets(2).Cells(2, 5).HorizontalAlignment = xlGeneral
    
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

    ' Surface characteristics
    MyBk.Sheets(2).Cells(2, 8).Value = "SURFACE CHARACTERISTICS"
    MyBk.Sheets(2).Cells(2, 8).HorizontalAlignment = xlGeneral
    MyBk.Sheets(2).Cells(3, 8).Value = "D65 (mm)"
    MyBk.Sheets(2).Cells(3, 9).Value = Format(InSt.Cells(13, 2).Value, "###0.0")
    MyBk.Sheets(2).Cells(4, 8).Value = "Gravel Fraction"
    MyBk.Sheets(2).Cells(4, 9).Value = Format(InSt.Cells(12, 1).Value, "0.##0")
    MyBk.Sheets(2).Cells(5, 8).Value = "Sand Fraction"
    MyBk.Sheets(2).Cells(5, 9).Value = Format(1 - InSt.Cells(12, 1).Value, "0.##0")

    ' Bedload Sampling
    Application.ScreenUpdating = False
    MyBk.Activate
    Sheets(2).Select
    MyBk.Sheets(2).Cells(7, 8).Value = "BEDLOAD SAMPLING"
    MyBk.Sheets(2).Cells(7, 8).HorizontalAlignment = xlGeneral
    MyBk.Sheets(2).Cells(8, 8).Value = "Qw (cms)"
    AddCommentsToCell Range("H8"), "water discharge"
    MyBk.Sheets(2).Cells(8, 9).Value = "QG (kg/min.)"
    AddCommentsToCell Range("I8"), "gravel transport rate"
    MyBk.Sheets(2).Cells(8, 10).Value = "QS (kg/min.)"
    AddCommentsToCell Range("J8"), "sand transport rate"
    Sheets(1).Select
    ThisWorkbook.Activate
    Application.ScreenUpdating = False
    i = 0
    Do While Not IsEmpty(InSt.Cells(i + 1, 18))
        i = i + 1
        MyBk.Sheets(2).Cells(i + 8, 8).Value = InSt.Cells(i, 18).Value
        QG(0) = InSt.Cells(i, 19).Value * InSt.Cells(i, 20).Value
        If TaurG > 0 And QG(0) > 0 Then _
            MyBk.Sheets(2).Cells(i + 8, 9).Value = QG(0)
        Qs(0) = InSt.Cells(i, 19).Value * (1 - InSt.Cells(i, 20).Value)
        If TaurS > 0 And Qs(0) > 0 Then _
            MyBk.Sheets(2).Cells(i + 8, 10).Value = Qs(0)
        If InSt.Cells(i, 40).Value Then _
            MyBk.Sheets(2).Cells(i + 8, 11).Value = "Outlier"
    Loop
    Nsample = i
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
    
    ' bedload transport rates
    If InSt.Cells(6, 1).Value = "(A)" Then 'single discharge
        MyBk.Sheets(3).Cells(2, 2).Value = "Gravel transport rate (kg/min.)"
        MyBk.Sheets(3).Cells(2, 2).HorizontalAlignment = xlGeneral
        MyBk.Sheets(3).Cells(2, 6).Value = QG(1) * 2650 * 60 ' m3/s to kg/min.
        MyBk.Sheets(3).Cells(3, 2).Value = "Sand transport rate (kg/min.)"
        MyBk.Sheets(3).Cells(3, 2).HorizontalAlignment = xlGeneral
        MyBk.Sheets(3).Cells(3, 6).Value = Qs(1) * 2650 * 60 ' m3/s to kg/min.
    Else
        If InSt.Cells(6, 1).Value = "(B)" Then 'min. and max. discharge
            MyBk.Sheets(3).Cells(2, 2).Value = "Rating curves are presented starting Column H"
        Else 'duration curve
            QG(0) = 0
            Qs(0) = 0
            For i = 1 To 25
                QG(0) = QG(0) + 0.5 * (QG(i) + QG(i + 1)) * _
                    Abs(InSt.Cells(i + 1, 17).Value - InSt.Cells(i, 17).Value) / 100
                Qs(0) = Qs(0) + 0.5 * (Qs(i) + Qs(i + 1)) * _
                    Abs(InSt.Cells(i + 1, 17).Value - InSt.Cells(i, 17).Value) / 100
            Next
            MyBk.Sheets(3).Cells(2, 2).Value = "Average gravel transport rate (kg/min.)"
            MyBk.Sheets(3).Cells(2, 2).HorizontalAlignment = xlGeneral
            MyBk.Sheets(3).Cells(3, 2).Value = "Average sand transport rate (kg/min.)"
            MyBk.Sheets(3).Cells(3, 2).HorizontalAlignment = xlGeneral
            MyBk.Sheets(3).Cells(2, 6).Value = QG(0) * 2650 * 60 ' m3/s to kg/min.
            MyBk.Sheets(3).Cells(3, 6).Value = Qs(0) * 2650 * 60 ' m3/s to kg/min.
        End If
        MyBk.Sheets(3).Cells(4, 8).Value = "RATING CURVES"
        MyBk.Sheets(3).Cells(5, 8).Value = "Discharge"
        MyBk.Sheets(3).Cells(6, 8).Value = "(cms)"
        MyBk.Sheets(3).Cells(5, 9).Value = "Gravel transport"
        MyBk.Sheets(3).Cells(6, 9).Value = "rate (kg/min.)"
        MyBk.Sheets(3).Cells(5, 10).Value = "Sand transport"
        MyBk.Sheets(3).Cells(6, 10).Value = "rate (kg/min)"
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
            If TaurG > 0 And QG(i) > 0 Then _
                MyBk.Sheets(3).Cells(i + 6, 9).Value = QG(i) * 2650 * 60
            If TaurS > 0 And Qs(i) > 0 Then _
                MyBk.Sheets(3).Cells(i + 6, 10).Value = Qs(i) * 2650 * 60
            MyBk.Sheets(3).Cells(i + 6, 11).Value = H(i)
            If InSt.Cells(1, 1).Value = "XS" Then
                GetRhAndAreaFromDepth H(i), Rh, Area, Rh1, Area1, Rh2, Area2
                MyBk.Sheets(3).Cells(i + 6, 12).Value = Rh
            End If
        Next
    End If
    
    If InSt.Cells(6, 1).Value = "(A)" Then GoTo 10 'Single discharge
    
    cc = MyBk.Sheets(3).Cells(5, 11).Value & " " & MyBk.Sheets(3).Cells(6, 11).Value
    
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
'------------
    
    MyBk.Activate
    Worksheets("Output").Select
    Set xRange = Range(MyBk.Worksheets("Output").Cells(7, 8), MyBk.Worksheets("Output").Cells(32, 8))
    Set yRange = Range(MyBk.Worksheets("Output").Cells(7, 9), MyBk.Worksheets("Output").Cells(32, 9))
    AddRatingCurves MyBk, "Output", xRange, yRange, "Plot Gravel", _
        "Discharge (cms)", "Gravel Transport Rate (kg/min.)"
    
    MyBk.Activate
    Worksheets("Output").Select
    Set yRange = Range(MyBk.Worksheets("Output").Cells(7, 10), MyBk.Worksheets("Output").Cells(32, 10))
    AddRatingCurves MyBk, "Output", xRange, yRange, "Plot Sand", _
        "Discharge (cms)", "Sand Transport Rate (kg/min.)"
    
    MyBk.Activate
    Worksheets("Output").Select
    Set xRange = Range(MyBk.Worksheets("Input").Cells(9, 8), MyBk.Worksheets("Input").Cells(Nsample + 8, 8))
    Set yRange = Range(MyBk.Worksheets("Input").Cells(9, 9), MyBk.Worksheets("Input").Cells(Nsample + 8, 9))
    AddExperimentalData MyBk, xRange, yRange, "Plot Gravel"
    
    Set yRange = Range(MyBk.Worksheets("Input").Cells(9, 10), MyBk.Worksheets("Input").Cells(Nsample + 8, 10))
    AddExperimentalData MyBk, xRange, yRange, "Plot Sand"
    
    MyBk.Activate
    Worksheets("Output").Select
    Set xRange = Range(MyBk.Worksheets("Output").Cells(7, 8), MyBk.Worksheets("Output").Cells(32, 8))
    Set yRange = Range(MyBk.Worksheets("Output").Cells(7, 11), MyBk.Worksheets("Output").Cells(32, 11))
    AddRatingCurves MyBk, "Output", xRange, yRange, "Plot Depth", "Discharge (cms)", ULCase(cc)
    ModifyYaxisToNormal MyBk, "Plot Depth"
    
10  MsgBox "Calculation results with the two-fraction equation of Wilcock (2001) " & _
        "are temporarily stored in workbook " & MyBk.Name & ".  Please save the file with an appropriate " & _
        "file name in an appropriate folder upon finishing of the rest of the run." & vbLf & vbLf & _
        "Click ""OK"" to continue.", vbOKOnly + vbInformation, "Wilcock (2001)"
    
End Sub

Sub AuthorCreateBedloadsamplingForWilcockTwoFractionModelTesting()
    Dim Pswd As String
    Dim Ncal As Long, i As Long, k As Long
    Dim InSt As Worksheet, ResSt As Workbook
    Dim Qs(27) As Double, QG(27) As Double, H(27) As Double
    Dim Qw(27) As Double, Qw1 As Double, Qw2 As Double, Qwc As Double
    Dim Width As Double, Slope As Double, Area As Double, Rh As Double
    Dim Rough As Double, TaurG As Double, TaurS As Double
    Dim Ustar As Double, tau As Double
    Dim Fg As Double, Fs As Double
    Dim Rho As Double, R As Double, g As Double
    
    Pswd = Application.InputBox("Enter password please:", "Author only")
    If Pswd <> "not4you" Then Exit Sub
    
    Set InSt = Worksheets("Input")
    
    Ncal = 26
    Rough = 2 * 45 / 1000 ' twice of D65 (45 mm), convert from mm to m
    Slope = 0.002
    Fg = 0.7
    Fs = 1 - Fg
    Rho = 1000
    R = 1.65
    g = 9.81
            
    TaurG = 0.04 * Rho * R * g * Rough / 2: TaurS = 0.1 * Rho * R * g * 0.5 / 1000
        
    Width = 50
    Qw(0) = 2
    For i = 1 To Ncal
        Qw(i) = Qw(i - 1) * 1.3
        H(i) = (Qw(i) / 8.1 / Width) ^ 0.6 * Rough ^ 0.1 / (g * Slope) ^ 0.3
        Ustar = (g * H(i) * Slope) ^ 0.5
        tau = Rho * Ustar ^ 2
        If tau > TaurG Then
            QG(i) = Fg * Width * Ustar ^ 3 / R / g * 11.2 * (1 - 0.846 * TaurG / tau) ^ 4.5
        Else
            QG(i) = Fg * Width * Ustar ^ 3 / R / g * 0.0025 * (tau / TaurG) ^ 14.2
        End If
        Qs(i) = Fs * Width * Ustar ^ 3 / R / g * 11.2 * (1 - 0.846 * (TaurS / tau) ^ 0.5) ^ 4.5
    Next 'i=1 to Ncal
    
    'add workbook with two worksheets
    i = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = 1
    Set ResSt = Workbooks.Add
    Application.SheetsInNewWorkbook = i
    ResSt.Sheets(1).Cells(1, 1).Value = "Qw (cms)"
    ResSt.Sheets(1).Cells(1, 2).Value = "QT (kg/min.)"
    ResSt.Sheets(1).Cells(1, 3).Value = "fG"
    For i = 1 To Ncal
        ResSt.Sheets(1).Cells(i + 1, 1).Value = Qw(i)
        ResSt.Sheets(1).Cells(i + 1, 2).Value = (QG(i) + Qs(i)) * 2650 * 60 * Exp(2.5 * (Rnd - 0.5))
        ResSt.Sheets(1).Cells(i + 1, 3).Value = QG(i) / (QG(i) + Qs(i))
        InSt.Cells(i, 18).Value = Qw(i)
        InSt.Cells(i, 19).Value = ResSt.Sheets(1).Cells(i + 1, 2).Value
        InSt.Cells(i, 20).Value = ResSt.Sheets(1).Cells(i + 1, 3).Value
    Next
    ResSt.Sheets(1).Cells(2, 5).Value = "Channel width = " & Width & " m"
    ResSt.Sheets(1).Cells(3, 5).Value = "Surface D65 = " & Rough / 2 * 1000 & " mm"
    ResSt.Sheets(1).Cells(4, 5).Value = "Surface gravel fraction = " & Fg
End Sub



