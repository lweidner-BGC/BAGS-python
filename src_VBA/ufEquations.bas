Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1
Option Explicit

Const RoughnessEnabled As Boolean = True

Private Sub cbAccept_Click()
    Dim i As Long, nn As Double
    Parker90 = cbParker90.Value
    Parker82 = cbParker82.Value
    PK82 = cbPK82.Value
    Wilcock = cbWilcock.Value
    Wilcock03 = cbWilcock03.Value
    Bakke = cbBakke.Value
    If Parker90 Or Parker82 Or PK82 Or Wilcock Or Wilcock03 Or Bakke Then
        Me.Hide
        Unload Me
        UserFormInUse = False
    Else
        MsgBox "No equation is selected!  You must select at least one equation " & _
            "in order to continue the calculation.  You can click ""Help"" to learn the " & _
            "input data requirement or click ""Cancel"" to quit " & _
            "the program.", vbOKOnly + vbCritical, "Select equations"
    End If
    Worksheets("Input").Cells(8, 2).Value = Parker90
    Worksheets("Input").Cells(9, 2).Value = Parker82
    Worksheets("Input").Cells(8, 1).Value = PK82
    Worksheets("Input").Cells(10, 2).Value = Wilcock
    Worksheets("Input").Cells(10, 1).Value = Wilcock03
    Worksheets("Input").Cells(11, 2).Value = Bakke
    'cells(9,1) and (11,1) are reserved for future use
    If RoughnessEnabled Then
        If Parker90 Or Parker82 Or PK82 Or Wilcock03 Then
            i = MsgBox("You have chosen one or more uncalibrated equations (i.e., Parker 1990, " & _
                "Parker, Klingeman and McLean 1982, Parker and Klingeman 1982, Wilcock and Crowe 2003)." & _
                vbLf & vbLf & _
                "If you have a good estimate of Manning's n value in the main channel, you can elect " & _
                "to correct the default roughness to the above equations by using your estimated main channel Manning's n value." & _
                vbLf & vbLf & _
                "Apply roughness correction?", vbYesNo + vbQuestion, "Roughness Correction")
            If i = vbYes Then
                If Parker90 Then
                    nn = Worksheets("Input").Cells(17, 2).Value
                ElseIf Parker82 Then
                    nn = Worksheets("Input").Cells(18, 2).Value
                ElseIf PK82 Then
                    nn = Worksheets("Input").Cells(19, 2).Value
                Else
                    nn = Worksheets("Input").Cells(20, 2).Value
                End If
                nn = Application.InputBox("Enter your estimate of Manning's n value in the main channel:", "Manning's n", nn)
                If Parker90 Then
                    Worksheets("Input").Cells(17, 1).Value = "TRUE"
                    Worksheets("Input").Cells(17, 2).Value = nn
                End If
                If Parker82 Then
                    Worksheets("Input").Cells(18, 1).Value = "TRUE"
                    Worksheets("Input").Cells(18, 2).Value = nn
                End If
                If PK82 Then
                    Worksheets("Input").Cells(19, 1).Value = "TRUE"
                    Worksheets("Input").Cells(19, 2).Value = nn
                End If
                If Wilcock03 Then
                    Worksheets("Input").Cells(20, 1).Value = "TRUE"
                    Worksheets("Input").Cells(20, 2).Value = nn
                End If
            Else
                For i = 17 To 20
                    Worksheets("Input").Cells(i, 1).Value = "FALSE"
                Next
            End If
        End If
    Else
        For i = 17 To 20
            Worksheets("Input").Cells(i, 1).Value = "FALSE"
        Next
    End If
End Sub

Private Sub cbCancel_Click()
    Parker90 = False
    Parker82 = False
    PK82 = False
    Wilcock = False
    Wilcock03 = False
    Bakke = False
    Me.Hide
    Unload Me
    UserFormInUse = False
End Sub

Private Sub cbHelp_Click()
    Me.Hide
    Load ufAvailableInput
    ufAvailableInput.Show
End Sub

Private Sub CommandButton1_Click()
    MsgBox "Parker, G., and Klingeman, P.C. (1982)  " & _
        "On why gravel bed streams are paved, Water Resources Research, 18(5), " & _
        "1409-1423." & vbLf & vbLf & _
        "Input dat requirement: " & _
        "typical cross section (or bankfull width), " & _
        "reach average water surface slope at a relatively high discharge " & _
        "(or bed slope as an approximation), and substrate grain size distribution. " & _
        "Flow duration curve is also needed if the user wish to calculate an average " & _
        "bedload transport rate.", vbOKOnly, "BAGS"
End Sub

Private Sub HelpBakke_Click()
    MsgBox "Bakke, P.D., Basdekas, P.O., Dawdy, D.R., and Klingeman, P.C. (1999) " & _
        "Calibrated Parker-Klingeman model for gravel transport, Journal of Hydraulic " & _
        "Engineering, 125(6), 657-660." & vbLf & vbLf & _
        "Input data requirement:  " & _
        "typical cross section (or bankfull width), " & _
        "reach average water surface slope at a relatively high discharge " & _
        "(or bed slope as an approximation), substrate grain size distribution, " & _
        "and bedload sampling results, including grain size distribution.  " & _
        "Flow duration curve is also needed if " & _
        "the user wish to calculate an average bedload transport rate.", _
        vbOKOnly, "BAGS"
End Sub

Private Sub HelpParker82_Click()
    MsgBox "Parker, G., Klingeman, P.C., and McLean, D.G. (1982) Bedload and size " & _
        "distribution in paved gravel-bed streams, Journal of Hydraulic Division, " & _
        "108(HY4), 544-571" & vbLf & vbLf & _
        "Input data requirement:  water discharge (or flow duration curve), " & _
        "typical cross section (or bankfull width), " & _
        "reach average water surface slope at a relatively high discharge " & _
        "(or bed slope as an approximation), substrate D50 or grain size distribution.", _
        vbOKOnly, "BAGS"
End Sub

Private Sub HelpParker90_Click()
    MsgBox "Parker, G. (1990) Surface-based bedload transport relation for gravel rivers, " & _
        "Journal of Hydraulic Research, 28(4), 417-430." & vbLf & vbLf & _
        "Input data requirement:  water discharge (or flow duration curve), " & _
        "typical cross section (or bankfull width), " & _
        "reach average water surface slope at a relatively high discharge " & _
        "(or bed slope as an approximation), surface grain size distribution.", _
        vbOKOnly, "BAGS"
End Sub

Private Sub HelpWilcock_Click()
    MsgBox "Wilcock, P.R., (2001) Toward a practical method for estimating sediment" & _
        "-transport rates in gravel-bed rivers, Earth Surface Processes and Landforms, " & _
        "26, 1395-1408." & vbLf & vbLf & _
        "Input data requirement:  water discharge (or flow duration curve), " & _
        "typical cross section (or bankfull width), " & _
        "reach average water surface slope at a relatively high discharge " & _
        "(or bed slope as an approximation), surface grain size distribution " & _
        "(or fractions of sand and gravel in channel bed and an estimate of " & _
        "surface D65), and a few bedload sampling results, including gravel/sand fractions.", _
        vbOKOnly, "BAGS"
End Sub

Private Sub HelpWilcock03_Click()
    MsgBox "Wilcock, P.R., and Crowe, J.C. (2003) Surface-based transport model " & _
        "for mixed-size sediment, Journal of Hydraulic Engineering, 129(2), 120-128." & _
        vbLf & vbLf & _
        "Input data requirement:  water discharge (or flow duration curve), " & _
        "typical cross section (or bankfull width), " & _
        "reach average water surface slope at a relatively high discharge " & _
        "(or bed slope as an approximation), surface grain size distribution.", _
        vbOKOnly, "BAGS"
End Sub


