Attribute VB_Name = "ChartCreator"
' ============================================================
' ChartCreator.bas
' Excel Canvas Chart Recreation Add-In
' Import this module into your VBA project, then run ShowChartCreator
' ============================================================

Option Explicit

' Supported chart type constants (maps to xlChartType)
Public Const CC_BAR         As Long = xlBarClustered          ' 57
Public Const CC_COLUMN      As Long = xlColumnClustered        ' 51
Public Const CC_LINE        As Long = xlLine                   ' 4
Public Const CC_LINE_MARK   As Long = xlLineMarkers            ' 65
Public Const CC_PIE         As Long = xlPie                    ' 5
Public Const CC_DOUGHNUT    As Long = xlDoughnut               ' -4120
Public Const CC_AREA        As Long = xlArea                   ' 1
Public Const CC_SCATTER     As Long = xlXYScatter              ' -4169
Public Const CC_SCATTER_LN  As Long = xlXYScatterLines         ' 74
Public Const CC_RADAR       As Long = xlRadar                  ' -4151
Public Const CC_STACKED_BAR As Long = xlBarStacked             ' 58
Public Const CC_STACKED_COL As Long = xlColumnStacked          ' 52

' ============================================================
' Entry point — assign this to a button or keyboard shortcut
' ============================================================
Public Sub ShowChartCreator()
    Dim frm As ChartForm
    Set frm = New ChartForm
    frm.Show
    Set frm = Nothing
End Sub

' ============================================================
' Core chart creation — called by the form
' ============================================================
Public Sub CreateChart( _
    dataRange   As Range, _
    chartType   As Long, _
    chartTitle  As String, _
    hasHeaders  As Boolean, _
    placement   As String, _
    showLegend  As Boolean, _
    showLabels  As Boolean, _
    colorScheme As String, _
    chartW      As Double, _
    chartH      As Double _
)

    Dim ws        As Worksheet
    Dim ch        As ChartObject
    Dim cht       As Chart
    Dim plotRange As Range

    Set ws = dataRange.Worksheet
    Set plotRange = dataRange

    ' ---- Place chart ----
    If placement = "NewSheet" Then
        Dim newSheet As Chart
        Set newSheet = Charts.Add()
        Set cht = newSheet
        cht.SetSourceData Source:=plotRange, PlotBy:=xlColumns
    Else
        Dim leftPos As Double
        Dim topPos  As Double

        If placement = "Right" Then
            leftPos = dataRange.Offset(0, dataRange.Columns.Count + 1).Left
            topPos  = dataRange.Top
        ElseIf placement = "Below" Then
            leftPos = dataRange.Left
            topPos  = dataRange.Offset(dataRange.Rows.Count + 1, 0).Top
        Else
            leftPos = dataRange.Offset(0, dataRange.Columns.Count + 1).Left
            topPos  = dataRange.Top
        End If

        Set ch = ws.ChartObjects.Add( _
            Left:=leftPos, Top:=topPos, Width:=chartW, Height:=chartH)
        Set cht = ch.Chart
        cht.SetSourceData Source:=plotRange, PlotBy:=xlColumns
    End If

    ' ---- Chart type ----
    cht.ChartType = chartType

    ' ---- Headers ----
    If hasHeaders Then
        On Error Resume Next
        cht.SeriesCollection(1).XValues = dataRange.Rows(1)
        On Error GoTo 0
    End If

    ' ---- Title ----
    If chartTitle <> "" Then
        cht.HasTitle = True
        cht.ChartTitle.Text = chartTitle
        With cht.ChartTitle.Font
            .Bold = False
            .Size = 11
            .Color = RGB(100, 100, 100)
        End With
    Else
        cht.HasTitle = False
    End If

    ' ---- Legend ----
    cht.HasLegend = showLegend
    If showLegend Then
        cht.Legend.Position = xlLegendPositionBottom
    End If

    ' ---- Data labels ----
    Dim s As Series
    For Each s In cht.SeriesCollection
        s.HasDataLabels = showLabels
        If showLabels Then
            With s.DataLabels
                .ShowValue = True
                .ShowSeriesName = False
                .ShowCategoryName = False
                .Font.Size = 9
            End With
        End If
    Next s

    ' ---- Colors ----
    ApplyColorScheme cht, colorScheme, chartType

    ' ---- Style (Canva = full editorial look; others = basic axis formatting) ----
    If colorScheme = "Canva" Then
        ApplyCanvaStyle cht, chartType, showLegend
    Else
        If chartType <> xlPie And chartType <> xlDoughnut And chartType <> xlRadar Then
            FormatAxes cht
        End If
    End If

    MsgBox "Chart created successfully!", vbInformation, "Chart Creator"

End Sub

' ============================================================
' CANVA STYLE — clean white editorial look matching Canva output
' ============================================================
Private Sub ApplyCanvaStyle(cht As Chart, chartType As Long, showLegend As Boolean)

    Const GRAY_LABEL  As Long = 8947848   ' RGB(136,136,136)
    Const GRAY_GRID   As Long = 14540253  ' RGB(221,221,221)
    Const WHITE       As Long = 16777215  ' RGB(255,255,255)

    ' ---- Chart area: white, no border ----
    With cht.ChartArea
        .Interior.Color = WHITE
        .Format.Line.Visible = msoFalse
        .RoundedCorners = False
    End With

    ' ---- Plot area: white, no border ----
    With cht.PlotArea
        .Interior.Color = WHITE
        .Format.Line.Visible = msoFalse
    End With

    ' ---- Legend: no border, small gray font ----
    If showLegend And cht.HasLegend Then
        With cht.Legend
            .Format.Line.Visible = msoFalse
            .Interior.Color = WHITE
            With .Font
                .Size  = 8
                .Color = GRAY_LABEL
                .Bold  = False
            End With
        End With
    End If

    ' ---- Title font ----
    If cht.HasTitle Then
        With cht.ChartTitle.Font
            .Size  = 11
            .Color = RGB(80, 80, 80)
            .Bold  = False
        End With
    End If

    ' ---- Axes ----
    On Error Resume Next

    Select Case chartType

        Case xlPie
            ' No axes — just clean up labels
            Dim sp As Series
            For Each sp In cht.SeriesCollection
                sp.HasDataLabels = True
                With sp.DataLabels
                    .ShowPercentage = True
                    .ShowValue      = False
                    .Font.Size      = 9
                    .Font.Color     = RGB(80, 80, 80)
                    .Position       = xlLabelPositionOutsideEnd
                End With
            Next sp

        Case xlDoughnut
            ' Doughnut: large hole, percentage labels on segments
            cht.SeriesCollection(1).DoughnutHoleSize = 60
            Dim sd As Series
            For Each sd In cht.SeriesCollection
                sd.HasDataLabels = True
                With sd.DataLabels
                    .ShowPercentage   = True
                    .ShowValue        = False
                    .Font.Size        = 9
                    .Font.Color       = RGB(255, 255, 255)
                    .Font.Bold        = False
                End With
            Next sd

        Case xlRadar
            ' Radar: style gridlines only
            Dim axR As Axis
            Set axR = cht.Axes(xlValue)
            With axR.MajorGridlines.Format.Line
                .Visible        = msoTrue
                .ForeColor.RGB  = GRAY_GRID
                .Weight         = 0.5
            End With

        Case Else
            ' Column, Bar, Line, Area, Scatter, Stacked variants
            Dim axCat As Axis
            Dim axVal As Axis
            Set axCat = cht.Axes(xlCategory)
            Set axVal = cht.Axes(xlValue)

            ' Category axis
            With axCat
                .HasTitle = False
                .Format.Line.Visible = msoFalse     ' no axis line
                .MajorTickMark = xlNone
                .MinorTickMark = xlNone
                With .TickLabels.Font
                    .Size  = 9
                    .Color = GRAY_LABEL
                    .Bold  = False
                End With
                ' No category gridlines
                If .HasMajorGridlines Then
                    .MajorGridlines.Format.Line.Visible = msoFalse
                End If
            End With

            ' Value axis
            With axVal
                .HasTitle = False
                .Format.Line.Visible = msoFalse     ' no axis line
                .MajorTickMark = xlNone
                .MinorTickMark = xlNone
                With .TickLabels.Font
                    .Size  = 9
                    .Color = GRAY_LABEL
                    .Bold  = False
                End With
                ' Light gray solid gridlines (Canva look)
                If Not .MajorGridlines Is Nothing Then
                    With .MajorGridlines.Format.Line
                        .Visible       = msoTrue
                        .ForeColor.RGB = GRAY_GRID
                        .Weight        = 0.5
                        .DashStyle     = msoLineSolid
                    End With
                End If
            End With

            ' For bar charts the axes are swapped
            If chartType = xlBarClustered Or chartType = xlBarStacked Then
                Dim axSer As Axis
                Set axSer = cht.Axes(xlSeriesAxis)
                If Not axSer Is Nothing Then
                    axSer.Format.Line.Visible = msoFalse
                End If
            End If

    End Select

    On Error GoTo 0

End Sub

' ============================================================
' Apply a named color scheme to all series
' ============================================================
Private Sub ApplyColorScheme(cht As Chart, scheme As String, chartType As Long)

    Dim palettes As Variant

    Select Case scheme

        ' ---- Canva editorial palette (matched from images) ----
        Case "Canva"
            If chartType = xlDoughnut Or chartType = xlPie Then
                ' 5-colour muted palette for doughnut/pie (matches image 3)
                palettes = Array( _
                    RGB(148, 58,  58),   ' dark red/burgundy  (<2.2)
                    RGB(52,  88,  92),   ' dark teal          (2.2-2.4)
                    RGB(183, 177, 165),  ' light stone        (2.4-2.6)
                    RGB(72,  83,  86),   ' dark slate         (2.6-2.8)
                    RGB(212, 204, 188))  ' light cream        (>2.8)
            Else
                ' 4-colour muted palette for column/line (matches images 1 & 2)
                palettes = Array( _
                    RGB(68,  103, 106),  ' muted teal         (Dollar General)
                    RGB(188, 178, 156),  ' light tan/beige    (Dollar Tree)
                    RGB(32,  85,  92),   ' dark teal          (Five Below)
                    RGB(142, 127, 80),   ' olive/khaki        (Dollarama)
                    RGB(148, 58,  58),   ' dark red           (extra series 5)
                    RGB(72,  83,  86),   ' dark slate         (extra series 6)
                    RGB(212, 204, 188),  ' light cream        (extra series 7)
                    RGB(183, 177, 165))  ' light stone        (extra series 8)
            End If

        Case "Office"
            palettes = Array( _
                RGB(68, 114, 196), RGB(237, 125, 49), RGB(165, 165, 165), _
                RGB(255, 192, 0),  RGB(91, 155, 213),  RGB(112, 173, 71), _
                RGB(38, 68, 120),  RGB(158, 72, 14))

        Case "Vivid"
            palettes = Array( _
                RGB(255, 87, 51),  RGB(51, 181, 229), RGB(255, 195, 0), _
                RGB(76, 187, 23),  RGB(142, 68, 173),  RGB(26, 188, 156), _
                RGB(231, 76, 60),  RGB(52, 152, 219))

        Case "Pastel"
            palettes = Array( _
                RGB(174, 214, 241), RGB(250, 215, 160), RGB(171, 235, 198), _
                RGB(245, 183, 177), RGB(215, 189, 226), RGB(249, 231, 159), _
                RGB(209, 236, 241), RGB(253, 212, 230))

        Case "Greyscale"
            palettes = Array( _
                RGB(50, 50, 50),   RGB(100, 100, 100), RGB(150, 150, 150), _
                RGB(190, 190, 190), RGB(220, 220, 220), RGB(30, 30, 30), _
                RGB(80, 80, 80),   RGB(130, 130, 130))

        Case "Dark"
            palettes = Array( _
                RGB(0, 173, 181),  RGB(255, 170, 51),  RGB(220, 80, 80), _
                RGB(100, 200, 120), RGB(160, 100, 220), RGB(240, 110, 160), _
                RGB(50, 150, 200), RGB(200, 200, 60))

        Case Else  ' "Default" — leave Excel's own colours
            Exit Sub

    End Select

    ' Apply colors to each series (wraps if more series than palette entries)
    Dim i As Integer
    Dim s As Series
    i = 0
    For Each s In cht.SeriesCollection
        Dim clr As Long
        clr = palettes(i Mod (UBound(palettes) + 1))
        s.Format.Fill.ForeColor.RGB  = clr
        s.Format.Line.ForeColor.RGB  = clr
        ' For line charts: thicker line weight, no markers (Canva look)
        If chartType = xlLine Or chartType = xlLineMarkers Or _
           chartType = xlXYScatterLines Then
            s.Format.Line.Weight = 1.5
            If chartType = xlLine Then s.MarkerStyle = xlMarkerStyleNone
        End If
        i = i + 1
    Next s

End Sub

' ============================================================
' Basic axis formatting for non-Canva schemes
' ============================================================
Private Sub FormatAxes(cht As Chart)

    On Error Resume Next

    Dim axCat As Axis
    Dim axVal As Axis

    Set axCat = cht.Axes(xlCategory)
    Set axVal = cht.Axes(xlValue)

    With axCat
        .HasTitle = False
        .TickLabels.Font.Size = 10
        .AxisBetweenCategories = True
        If .HasMajorGridlines Then
            .MajorGridlines.Format.Line.Visible = msoFalse
        End If
    End With

    With axVal
        .HasTitle = False
        .TickLabels.Font.Size = 10
        With .MajorGridlines.Format.Line
            .Visible       = msoTrue
            .ForeColor.RGB = RGB(200, 200, 200)
            .DashStyle     = msoLineDash
        End With
    End With

    On Error GoTo 0

End Sub

' ============================================================
' Helper: return chart type Long from display name string
' ============================================================
Public Function ChartTypeFromName(name As String) As Long
    Select Case name
        Case "Bar (Horizontal)":        ChartTypeFromName = xlBarClustered
        Case "Column (Vertical)":       ChartTypeFromName = xlColumnClustered
        Case "Line":                    ChartTypeFromName = xlLine
        Case "Line with Markers":       ChartTypeFromName = xlLineMarkers
        Case "Pie":                     ChartTypeFromName = xlPie
        Case "Doughnut":                ChartTypeFromName = xlDoughnut
        Case "Area":                    ChartTypeFromName = xlArea
        Case "Scatter":                 ChartTypeFromName = xlXYScatter
        Case "Scatter with Lines":      ChartTypeFromName = xlXYScatterLines
        Case "Radar":                   ChartTypeFromName = xlRadar
        Case "Stacked Bar":             ChartTypeFromName = xlBarStacked
        Case "Stacked Column":          ChartTypeFromName = xlColumnStacked
        Case Else:                      ChartTypeFromName = xlColumnClustered
    End Select
End Function
