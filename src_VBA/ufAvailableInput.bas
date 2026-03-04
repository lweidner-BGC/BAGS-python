Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1
Option Explicit

Private Sub cbCancel_Click()
    Me.Hide
    Unload Me
    ufEquations.Show
End Sub

Private Sub cbOK_Click()
    Dim Parker90OK As Boolean
    Dim Parker82OK As Boolean
    Dim PK82OK As Boolean
    Dim WilcockOK As Boolean
    Dim Wilcock03OK As Boolean
    Dim BakkeOK As Boolean
    Dim NoneOK As Boolean
    Dim ExitSub As Label
    Dim MyMsg As String
    
    Parker90OK = False
    Parker82OK = False
    PK82OK = False
    WilcockOK = False
    Wilcock03OK = False
    BakkeOK = False
    NoneOK = False
    
    If Not cbCrossSection Then
        If Not cbWidth Then
            NoneOK = True
            GoTo ExitSub
        End If
    End If
    If Not cbWsSlope Then
        If Not cbBedSlope Then
            NoneOK = True
            GoTo ExitSub
        End If
    End If
    
    If cbSurfSize Then
        Parker90OK = True
        Wilcock03OK = True
    End If
    
    If cbSubSize Or cbDsub50 Then _
        Parker82OK = True
    
    If cbSubSize Then
        PK82OK = True
        If cbSampling And cbSamplingSize Then
            BakkeOK = True
        End If
    End If
    
    If (cbSurfSize Or (cbSGfraction And cbD65)) And _
        (cbSampling Or cbSamplingSize Or Me.cbSampleGSFraction) Then _
        WilcockOK = True
    
ExitSub:
    If NoneOK Then
        MyMsg = "You do not have enough input data to carry out a bedload transport calculation."
    Else
        MyMsg = "You can use the following equation(s) to carry out the bedload transport calculation:" & vbLf
        If Parker90OK Then
            MyMsg = MyMsg & vbLf & _
                "    Parker (surface-based, 1990)"
        End If
        If Parker82OK Then
            MyMsg = MyMsg & vbLf & _
                "    Parker, Klingeman, and McLean (substrate D50-based, 1982)"
        End If
        If PK82OK Then
            MyMsg = MyMsg & vbLf & _
                "    Parker and Klingeman (substrate-based, 1982)"
        End If
        If WilcockOK Then
            MyMsg = MyMsg & vbLf & _
                "    Wilcock (2001)"
        End If
        If Wilcock03OK Then
            MyMsg = MyMsg & vbLf & _
                "    Wilcock and Crowe (2003)"
        End If
        If BakkeOK Then
            MyMsg = MyMsg & vbLf & _
                "    Bakke et al. (1999)"
        End If
    End If
    MsgBox MyMsg, vbOKOnly + vbInformation, "Available equations"
    
    Me.Hide
    Unload Me
    ufEquations.Show
End Sub


