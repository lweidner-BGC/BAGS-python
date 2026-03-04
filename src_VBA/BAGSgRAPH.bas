Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

Const GraphicsEnabled As Boolean = True

Sub AddRatingCurves(MyBk As Workbook, MySource As String, xRange As Range, yRange As Range, _
    MyTab As String, MyXTitle As String, MyYTitle As String)
    Dim LogMaxLoad As Long, LogMinLoad As Long
    Dim LogMaxQw As Long, LogMinQw As Long
    Dim MaxLoad As Double, MinLoad As Double
    Dim MaxQw As Double, MinQw As Double
    Dim dmy As Double, i As Long
    
    If Not GraphicsEnabled Then Exit Sub
    
    On Error Resume Next
    
    If Application.WorksheetFunction.Min(yRange) > 0 Then
        LogMaxLoad = Int(Log(Application.WorksheetFunction.Max(yRange)) / Log(10) + 1)
        LogMinLoad = Int(Log(Application.WorksheetFunction.Min(yRange)) / Log(10))
        
        If LogMinLoad < LogMaxLoad - 7 Then
            LogMinLoad = LogMaxLoad - 7
        End If
        
        MaxLoad = 10 ^ LogMaxLoad
        MinLoad = 10 ^ LogMinLoad
    Else
        MaxLoad = 100
        MinLoad = 1
    End If
    
    If Application.WorksheetFunction.Min(xRange) > 0 Then
        LogMaxQw = Int(Log(Application.WorksheetFunction.Max(xRange)) / Log(10) + 1)
        MaxQw = 10 ^ LogMaxQw
        
        If MinLoad <= Application.WorksheetFunction.Min(yRange) Or _
                        Application.WorksheetFunction.Min(yRange) <= 0 Then
            LogMinQw = Int(Log(Application.WorksheetFunction.Min(xRange)) / Log(10))
        Else
            For i = 1 To xRange.Cells.Count - 1
                If (MinLoad >= yRange.Cells(i).Value And MinLoad <= yRange.Cells(i + 1).Value) Or _
                    (MinLoad <= yRange.Cells(i).Value And MinLoad >= yRange.Cells(i + 1).Value) Then
                    dmy = Log(xRange.Cells(i).Value) + (Log(xRange.Cells(i + 1).Value) - Log(xRange.Cells(i).Value)) / _
                        (Log(yRange.Cells(i + 1).Value) - Log(yRange.Cells(i).Value)) * _
                        (Log(MinLoad) - Log(yRange.Cells(i).Value))
                    dmy = Exp(dmy)
                    LogMinQw = Int(Log(dmy) / Log(10))
                    Exit For
                End If
            Next
        End If
        MinQw = 10 ^ LogMinQw
    Else
        MinQw = 0.1
        MaxQw = 100
    End If
    
    Charts.Add
    ActiveChart.ChartType = xlXYScatterLines
    ActiveChart.SetSourceData Source:=Sheets(MySource).Range(Application.Union(xRange, yRange).Address), _
        PlotBy:=xlColumns
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:=MyTab
    With ActiveChart
        .HasTitle = False
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = MyXTitle
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = MyYTitle
        .HasLegend = False
    End With
    With ActiveChart.Axes(xlCategory)
        .HasMajorGridlines = True
        .HasMinorGridlines = True
        .MinimumScale = MinQw
        .MaximumScale = MaxQw
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlCustom
        .CrossesAt = MinQw
        .ReversePlotOrder = False
        .ScaleType = xlLogarithmic
        .DisplayUnit = xlNone
    End With
    With ActiveChart.Axes(xlValue)
        .HasMajorGridlines = True
        .HasMinorGridlines = True
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .MinimumScale = MinLoad
        .MaximumScale = MaxLoad
        .Crosses = xlCustom
        .CrossesAt = MinLoad
        .ReversePlotOrder = False
        .ScaleType = xlLogarithmic
        .DisplayUnit = xlNone
    End With
    With ActiveChart.PlotArea.Border
        .ColorIndex = 1
        .Weight = xlThin
        .LineStyle = xlContinuous
    End With
    ActiveChart.PlotArea.Interior.ColorIndex = xlNone
    With ActiveChart.SeriesCollection(1).Border
        .Weight = xlMedium
        .LineStyle = xlContinuous
    End With
    ActiveChart.SeriesCollection(1).MarkerStyle = xlNone
    With ActiveChart.ChartArea.Font
        .Name = "Arial"
        .Size = 18
    End With
    ActiveChart.Axes(xlValue).AxisTitle.Font.Bold = True
    ActiveChart.Axes(xlCategory).AxisTitle.Font.Bold = True
    
    ActiveChart.Deselect
    Worksheets("Note").Select
    ThisWorkbook.Activate
End Sub

Sub AddExperimentalData(MyBk As Workbook, xRange As Range, yRange As Range, MyTab As String)
    MyBk.Activate
    
    If Not GraphicsEnabled Then Exit Sub
    
    Application.Union(xRange, yRange).Copy
    Sheets(MyTab).Select
    ActiveChart.SeriesCollection.Paste Rowcol:=xlColumns, SeriesLabels:=False, _
        CategoryLabels:=True, Replace:=False, NewSeries:=True
    With ActiveChart.SeriesCollection(2).Border
        .LineStyle = xlNone
    End With
    With ActiveChart.SeriesCollection(2)
        .MarkerStyle = xlCircle
        .MarkerBackgroundColorIndex = 2
        .MarkerForegroundColorIndex = 1
        .MarkerSize = 7
    End With

    ActiveChart.Deselect
    Worksheets("Note").Select
    ThisWorkbook.Activate
End Sub

Sub ModifyYaxisToNormal(MyBk As Workbook, MyTab As String)
    MyBk.Activate
    
    If Not GraphicsEnabled Then Exit Sub
    
    Charts(MyTab).Select
    
    With ActiveChart.Axes(xlValue)
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlCustom
        .CrossesAt = 0.1
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    ActiveChart.Axes(xlValue).CrossesAt = 0
    ActiveChart.Axes(xlValue).MinorGridlines.Delete
    
    ActiveChart.Deselect
    Worksheets("Note").Select
    ThisWorkbook.Activate
End Sub

Sub AdjustYaxisScale(MyBk As Workbook, MyChart As String, MinY As Long, MaxY As Long, MUnit As Long)
    MyBk.Activate
    
    If Not GraphicsEnabled Then Exit Sub
    
    Charts(MyChart).Select
    With ActiveChart.Axes(xlValue)
        .MinimumScale = 0
        .MaximumScale = 100
        .MajorUnit = 20
    End With
    ActiveChart.Deselect
    Worksheets("Note").Select
    ThisWorkbook.Activate
End Sub

Sub AdjustXaxisToNormal(MyBk As Workbook, MyTab As String)
    MyBk.Activate
    
    If Not GraphicsEnabled Then Exit Sub
    
    Charts(MyTab).Select
    
    With ActiveChart.Axes(xlCategory)
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    ActiveChart.Axes(xlCategory).MinorGridlines.Delete
        ActiveChart.Axes(xlCategory).CrossesAt = -200000
    
    ActiveChart.Deselect
    Worksheets("Note").Select
    ThisWorkbook.Activate
End Sub



