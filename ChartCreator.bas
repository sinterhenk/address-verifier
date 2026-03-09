Attribute VB_Name = "ChartCreator"
' ============================================================
' ChartCreator.bas  -  Excel Canvas Chart Recreation Add-In
'
' SINGLE FILE - just import this .bas into your VBA project.
' No .frm or .frx files needed.
'
' Run:  Alt+F8 > ShowChartCreator
' ============================================================

Option Explicit

' ============================================================
' ENTRY POINT - assign to a button / ribbon / keyboard shortcut
' ============================================================
Public Sub ShowChartCreator()
    ShowChartCreatorSimple
End Sub

' ============================================================
' Recolor series on the currently selected chart.
' Click a chart, then Alt+F8 > CustomizeChartColors
' ============================================================
Public Sub CustomizeChartColors()

    If ActiveChart Is Nothing Then
        MsgBox "Please click on a chart first, then run this macro.", _
               vbExclamation, "No chart selected"
        Exit Sub
    End If

    Dim cht As Chart
    Set cht = ActiveChart

    Dim s       As Series
    Dim i       As Integer
    Dim inp     As String
    Dim clr     As Long
    Dim msg     As String

    i = 1
    For Each s In cht.SeriesCollection

        Dim curHex As String
        curHex = ColorToHex(s.Format.Fill.ForeColor.RGB)

        msg = "Series " & i & ": " & s.Name & vbCrLf & vbCrLf & _
              "Current color: " & curHex & vbCrLf & vbCrLf & _
              "Enter new color as:" & vbCrLf & _
              "  Hex:  #RRGGBB  (e.g. #2A556C)" & vbCrLf & _
              "  RGB:  R,G,B    (e.g. 42,85,108)" & vbCrLf & vbCrLf & _
              "Leave blank to keep current color." & vbCrLf & _
              "(Cancel stops editing)"

        inp = InputBox(msg, "Series " & i & " of " & _
                       cht.SeriesCollection.Count & " - Customize Color", curHex)

        If StrPtr(inp) = 0 Then Exit For       ' Cancel
        If Trim(inp) = "" Then
            i = i + 1
        Else
            clr = ParseColor(Trim(inp))
            If clr = -1 Then
                MsgBox "Could not parse """ & inp & """." & vbCrLf & _
                       "Use #RRGGBB or R,G,B format.", vbExclamation, "Bad color"
            Else
                s.Format.Fill.ForeColor.RGB = clr
                s.Format.Line.ForeColor.RGB = clr
                i = i + 1
            End If
        End If
    Next s

    MsgBox "Colors updated!", vbInformation, "Chart Creator"
End Sub

' ============================================================
' InputBox-based workflow
' ============================================================
Private Sub ShowChartCreatorSimple()

    ' 1. Range
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox( _
        Prompt:="Select the data range for your chart:", _
        Title:="Chart Creator - Step 1 of 5: Data Range", _
        Default:=Selection.Address, _
        Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    ' 2. Chart type
    Dim typeList As String
    typeList = "1  Column (Vertical)" & vbCrLf & _
               "2  Bar (Horizontal)" & vbCrLf & _
               "3  Line" & vbCrLf & _
               "4  Line with Markers" & vbCrLf & _
               "5  Pie" & vbCrLf & _
               "6  Doughnut" & vbCrLf & _
               "7  Area" & vbCrLf & _
               "8  Scatter" & vbCrLf & _
               "9  Scatter with Lines" & vbCrLf & _
               "10 Stacked Column" & vbCrLf & _
               "11 Stacked Bar" & vbCrLf & _
               "12 Radar"

    Dim typeNum As String
    typeNum = InputBox("Enter the number of the chart type:" & vbCrLf & vbCrLf & _
                       typeList, "Chart Creator - Step 2 of 5: Chart Type", "1")
    If StrPtr(typeNum) = 0 Then Exit Sub
    If typeNum = "" Then typeNum = "1"

    Dim typeNames As Variant
    typeNames = Array("Column (Vertical)", "Bar (Horizontal)", "Line", _
                      "Line with Markers", "Pie", "Doughnut", "Area", _
                      "Scatter", "Scatter with Lines", "Stacked Column", _
                      "Stacked Bar", "Radar")
    Dim idx As Integer
    idx = CInt(typeNum) - 1
    If idx < 0 Or idx > 11 Then idx = 0
    Dim chartType As Long
    chartType = ChartTypeFromName(CStr(typeNames(idx)))

    ' 3. Data orientation (rows vs columns)
    Dim orientInput As String
    orientInput = InputBox( _
        "How is your data arranged?" & vbCrLf & vbCrLf & _
        "1  Each ROW is one series" & vbCrLf & _
        "   (e.g. Row 1 = Company A, Row 2 = Company B...)" & vbCrLf & _
        "   Categories (e.g. FY years) run across columns" & vbCrLf & vbCrLf & _
        "2  Each COLUMN is one series" & vbCrLf & _
        "   (e.g. Col A = Company A, Col B = Company B...)" & vbCrLf & _
        "   Categories (e.g. FY years) run down rows", _
        "Chart Creator - Step 3 of 10: Data Orientation", "2")
    If StrPtr(orientInput) = 0 Then Exit Sub
    If orientInput = "" Then orientInput = "2"
    Dim plotByRows As Boolean
    plotByRows = (Trim(orientInput) = "1")

    ' 4. Title
    Dim chartTitle As String
    chartTitle = InputBox("Enter a chart title (or leave blank):", _
                          "Chart Creator - Step 4 of 10: Title", "")
    If StrPtr(chartTitle) = 0 Then Exit Sub

    ' 4. Color scheme
    Dim colorInput As String
    colorInput = InputBox( _
        "Enter the number of the color scheme:" & vbCrLf & vbCrLf & _
        "1  Canva (muted, editorial)" & vbCrLf & _
        "2  Default (Excel colours)" & vbCrLf & _
        "3  Office" & vbCrLf & _
        "4  Vivid" & vbCrLf & _
        "5  Pastel" & vbCrLf & _
        "6  Greyscale" & vbCrLf & _
        "7  Dark", _
        "Chart Creator - Step 5 of 10: Colors", "1")
    If StrPtr(colorInput) = 0 Then Exit Sub
    If colorInput = "" Then colorInput = "1"

    Dim schemeNames As Variant
    schemeNames = Array("Canva", "Default", "Office", "Vivid", "Pastel", "Greyscale", "Dark")
    Dim sIdx As Integer
    sIdx = CInt(colorInput) - 1
    If sIdx < 0 Or sIdx > 6 Then sIdx = 0
    Dim colorScheme As String
    colorScheme = CStr(schemeNames(sIdx))

    ' 5. Placement
    Dim placeInput As String
    placeInput = InputBox( _
        "Where should the chart go?" & vbCrLf & vbCrLf & _
        "1  Right of data" & vbCrLf & _
        "2  Below data" & vbCrLf & _
        "3  New sheet", _
        "Chart Creator - Step 6 of 10: Placement", "1")
    If StrPtr(placeInput) = 0 Then Exit Sub
    If placeInput = "" Then placeInput = "1"

    Dim placement As String
    Select Case placeInput
        Case "2": placement = "Below"
        Case "3": placement = "NewSheet"
        Case Else: placement = "Right"
    End Select

    ' 6. Font
    Dim fontInput As String
    fontInput = InputBox( _
        "Enter the number of the font:" & vbCrLf & vbCrLf & _
        "1  Libre Baskerville  (Canva default - must be installed)" & vbCrLf & _
        "2  Calibri  (Excel default)" & vbCrLf & _
        "3  Arial" & vbCrLf & _
        "4  Segoe UI" & vbCrLf & _
        "5  Georgia  (serif)" & vbCrLf & _
        "6  Trebuchet MS" & vbCrLf & _
        vbCrLf & "Or type any font name directly.", _
        "Chart Creator - Step 7 of 10: Font", "1")
    If StrPtr(fontInput) = 0 Then Exit Sub
    If fontInput = "" Then fontInput = "1"

    Dim fontName As String
    Select Case Trim(fontInput)
        Case "1": fontName = "Libre Baskerville"
        Case "2": fontName = "Calibri"
        Case "3": fontName = "Arial"
        Case "4": fontName = "Segoe UI"
        Case "5": fontName = "Georgia"
        Case "6": fontName = "Trebuchet MS"
        Case Else: fontName = Trim(fontInput)
    End Select

    ' 7. Line weight (shown for all types; mainly affects line/scatter charts)
    Dim lwInput As String
    lwInput = InputBox( _
        "Line weight for line/scatter charts (in points):" & vbCrLf & vbCrLf & _
        "1  = thin" & vbCrLf & _
        "2  = medium (Canva default)" & vbCrLf & _
        "3  = thick" & vbCrLf & vbCrLf & _
        "Or type any number (e.g. 1.5, 2.5)", _
        "Chart Creator - Step 8 of 10: Line Weight", "2")
    If StrPtr(lwInput) = 0 Then Exit Sub
    If lwInput = "" Or Not IsNumeric(lwInput) Then lwInput = "2"

    Dim lineWeight As Double
    lineWeight = CDbl(lwInput)

    ' 8. Y-axis number format (optional)
    Dim yFmtInput As String
    yFmtInput = InputBox( _
        "Y-axis number format (leave blank for auto):" & vbCrLf & vbCrLf & _
        "Examples:" & vbCrLf & _
        "  0%      → 10%  (whole, default for % data)" & vbCrLf & _
        "  0.0%    → 10.1%" & vbCrLf & _
        "  0.00%   → 10.12%" & vbCrLf & _
        "  #,##0   → 1,234" & vbCrLf & _
        "  #,##0.0 → 1,234.5" & vbCrLf & vbCrLf & _
        "Leave blank to keep source data format (or use auto % rounding).", _
        "Chart Creator - Step 9 of 10: Y-Axis Format", "")
    If StrPtr(yFmtInput) = 0 Then Exit Sub

    ' 10. Y-axis gridline count
    Dim gridInput As String
    gridInput = InputBox( _
        "Number of Y-axis gridlines (horizontal lines):" & vbCrLf & vbCrLf & _
        "Examples (for 0%–50% data):" & vbCrLf & _
        "  5  → lines at 10%, 20%, 30%, 40%, 50% (10% steps)" & vbCrLf & _
        "  10 → lines at 5%, 10%, 15%... (5% steps)" & vbCrLf & _
        "  4  → lines at 12.5%, 25%, 37.5%, 50%" & vbCrLf & vbCrLf & _
        "Leave blank to let Excel decide automatically.", _
        "Chart Creator - Step 10 of 11: Y-Axis Gridlines", "")
    If StrPtr(gridInput) = 0 Then Exit Sub

    Dim yAxisLines As Integer
    If IsNumeric(Trim(gridInput)) And Trim(gridInput) <> "" Then
        yAxisLines = CInt(Trim(gridInput))
        If yAxisLines < 1 Then yAxisLines = 0
    Else
        yAxisLines = 0
    End If

    ' 11. Legend columns
    Dim legColInput As String
    legColInput = InputBox( _
        "Choose a legend layout (enter a number):" & vbCrLf & vbCrLf & _
        "  1  →  Column  – entries stacked in one column  (default)" & vbCrLf & _
        "  2  →  2 per row  – e.g. 4 series → 2×2 grid" & vbCrLf & _
        "  3  →  3 per row  – e.g. 6 series → 3×2 grid" & vbCrLf & vbCrLf & _
        "Any other number = that many entries per row." & vbCrLf & _
        "Leave blank to let Excel decide automatically.", _
        "Chart Creator - Step 11 of 11: Legend Layout", "1")
    If StrPtr(legColInput) = 0 Then Exit Sub

    Dim legendCols As Integer
    If IsNumeric(Trim(legColInput)) And Trim(legColInput) <> "" Then
        legendCols = CInt(Trim(legColInput))
        If legendCols < 1 Then legendCols = 0
    Else
        legendCols = 0
    End If

    ' 12. X-axis date label format (optional)
    Dim xFmtInput As String
    xFmtInput = InputBox( _
        "X-axis date label format (leave blank to keep as-is):" & vbCrLf & vbCrLf & _
        "Presets — type the number:" & vbCrLf & _
        "  1  FY18, FY19...      (fiscal/calendar year)" & vbCrLf & _
        "  2  Jan 18, May 18...  (month + short year)" & vbCrLf & _
        "  3  January 2018...    (full month + full year)" & vbCrLf & _
        "  4  2018, 2019...      (4-digit year only)" & vbCrLf & _
        "  5  Q1 18, Q2 18...    (quarter + short year)" & vbCrLf & vbCrLf & _
        "Or type any Excel date format directly, e.g.  yy  or  mmm-yy", _
        "Chart Creator - Step 12 of 12: X-Axis Date Format", "")
    If StrPtr(xFmtInput) = 0 Then Exit Sub

    Dim xAxisNumFmt As String
    Select Case Trim(xFmtInput)
        Case "1": xAxisNumFmt = """FY""yy"
        Case "2": xAxisNumFmt = "mmm yy"
        Case "3": xAxisNumFmt = "mmmm yyyy"
        Case "4": xAxisNumFmt = "yyyy"
        Case "5": xAxisNumFmt = """Q""q yy"
        Case "":  xAxisNumFmt = ""
        Case Else: xAxisNumFmt = Trim(xFmtInput)
    End Select

    CreateChart rng, chartType, chartTitle, True, placement, True, False, _
                colorScheme, 480, 300, fontName, lineWeight, Trim(yFmtInput), _
                plotByRows, yAxisLines, legendCols, xAxisNumFmt

End Sub

' ============================================================
' Core chart creation
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
    chartH      As Double, _
    Optional fontName      As String = "Libre Baskerville", _
    Optional lineWeight    As Double = 2, _
    Optional yAxisNumFmt   As String = "", _
    Optional plotByRows    As Boolean = False, _
    Optional yAxisLines    As Integer = 0, _
    Optional legendCols    As Integer = 0, _
    Optional xAxisNumFmt   As String = "" _
)

    Dim ws        As Worksheet
    Dim ch        As ChartObject
    Dim cht       As Chart
    Dim plotRange As Range

    Set ws = dataRange.Worksheet
    Set plotRange = dataRange

    If placement = "NewSheet" Then
        Dim newSheet As Chart
        Set newSheet = Charts.Add()
        Set cht = newSheet
        cht.SetSourceData Source:=plotRange, PlotBy:=IIf(plotByRows, xlRows, xlColumns)
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
        cht.SetSourceData Source:=plotRange, PlotBy:=IIf(plotByRows, xlRows, xlColumns)
    End If

    cht.ChartType = chartType

    If chartTitle <> "" Then
        cht.HasTitle = True
        cht.ChartTitle.Text = chartTitle
        With cht.ChartTitle.Font
            .Name = fontName
            .Bold = False
            .Size = 13
            .Color = RGB(80, 80, 80)
        End With
    Else
        cht.HasTitle = False
    End If

    cht.HasLegend = showLegend
    If showLegend Then
        cht.Legend.Position = xlLegendPositionBottom
    End If

    Dim s As Series
    For Each s In cht.SeriesCollection
        s.HasDataLabels = showLabels
        If showLabels Then
            With s.DataLabels
                .ShowValue = True
                .ShowSeriesName = False
                .ShowCategoryName = False
                .Font.Name = fontName
                .Font.Size = 15.6
            End With
        End If
    Next s

    ApplyColorScheme cht, colorScheme, chartType, lineWeight

    If colorScheme = "Canva" Then
        ApplyCanvaStyle cht, chartType, showLegend, fontName, yAxisNumFmt, yAxisLines, xAxisNumFmt, legendCols
    Else
        If chartType <> xlPie And chartType <> xlDoughnut And chartType <> xlRadar Then
            FormatAxes cht, fontName, yAxisNumFmt, yAxisLines, xAxisNumFmt
        End If
    End If

    MsgBox "Chart created!" & vbCrLf & vbCrLf & _
           "Tip: to recolor series, click the chart then" & vbCrLf & _
           "Alt+F8 > CustomizeChartColors.", _
           vbInformation, "Chart Creator"

End Sub

' ============================================================
' CANVA STYLE
' ============================================================
Private Sub ApplyCanvaStyle(cht As Chart, chartType As Long, showLegend As Boolean, fontName As String, _
                            Optional yAxisNumFmt As String = "", _
                            Optional yAxisLines As Integer = 0, _
                            Optional xAxisNumFmt As String = "", _
                            Optional legendCols As Integer = 0)

    Const GRAY_LABEL  As Long = 3355443   ' #333333
    Const GRAY_GRID   As Long = 14540253
    Const WHITE       As Long = 16777215

    With cht.ChartArea
        .Interior.Color = WHITE
        .Format.Line.Visible = msoFalse
        .RoundedCorners = False
    End With

    With cht.PlotArea
        .Interior.Color = WHITE
        .Format.Line.Visible = msoFalse
    End With

    If showLegend And cht.HasLegend Then
        With cht.Legend
            .Format.Line.Visible = msoFalse
            .Interior.Color = WHITE
            With .Font
                .Name = fontName: .Size = 12.1: .Color = GRAY_LABEL: .Bold = False
            End With
            ' Apply column layout now that font is finalised
            If legendCols > 0 Then
                Dim nS As Integer: nS = cht.SeriesCollection.Count
                If nS > legendCols Then
                    ' Force full-width layout first so entry widths are untruncated
                    .Width = cht.ChartArea.Width
                    Dim le As LegendEntry, maxW As Double, entryH As Double
                    maxW = 0: entryH = 0
                    For Each le In .LegendEntries
                        If le.Width  > maxW  Then maxW  = le.Width
                        If le.Height > entryH Then entryH = le.Height
                    Next le
                    Dim nR As Integer
                    nR = Int((nS + legendCols - 1) / legendCols)
                    ' Width: use measured entry width; fallback to chart proportion
                    If maxW > 0 Then
                        .Width = maxW * legendCols
                    Else
                        .Width = cht.ChartArea.Width * legendCols / nS
                    End If
                    ' Height: use measured entry height; fallback to font-based estimate
                    If entryH > 0 Then
                        .Height = entryH * nR * 1.8
                    Else
                        .Height = 12.1 * 2.4 * nR
                    End If
                End If
            End If
        End With
    End If

    If cht.HasTitle Then
        With cht.ChartTitle.Font
            .Name = fontName: .Size = 13: .Color = RGB(80, 80, 80): .Bold = False
        End With
    End If

    On Error Resume Next

    Select Case chartType

        Case xlPie
            Dim sp As Series
            For Each sp In cht.SeriesCollection
                sp.HasDataLabels = True
                With sp.DataLabels
                    .ShowPercentage = True: .ShowValue = False
                    .NumberFormat = "0.0%"
                    .Font.Name = fontName: .Font.Size = 15.6: .Font.Color = RGB(80, 80, 80)
                    .Position = xlLabelPositionOutsideEnd
                End With
            Next sp

        Case xlDoughnut
            cht.SeriesCollection(1).DoughnutHoleSize = 60
            Dim sd As Series
            For Each sd In cht.SeriesCollection
                sd.HasDataLabels = True
                With sd.DataLabels
                    .ShowPercentage = True: .ShowValue = False
                    .NumberFormat = "0.0%"
                    .Font.Name = fontName: .Font.Size = 15.6
                    .Font.Color = RGB(255, 255, 255): .Font.Bold = False
                End With
            Next sd

        Case xlRadar
            Dim axR As Axis
            Set axR = cht.Axes(xlValue)
            With axR.MajorGridlines.Format.Line
                .Visible = msoTrue: .ForeColor.RGB = GRAY_GRID: .Weight = 0.5
            End With
            axR.TickLabels.Font.Name = fontName

        Case Else
            Dim axCat As Axis, axVal As Axis
            Set axCat = cht.Axes(xlCategory)
            Set axVal = cht.Axes(xlValue)

            With axCat
                .HasTitle = False
                .Format.Line.Visible = msoFalse
                .MajorTickMark = xlNone: .MinorTickMark = xlNone
                With .TickLabels.Font
                    .Name = fontName: .Size = 13.1: .Color = GRAY_LABEL: .Bold = False
                End With
                If xAxisNumFmt <> "" Then
                    On Error Resume Next
                    .AxisType = xlTimeScale
                    On Error GoTo 0
                    .TickLabels.NumberFormat = xAxisNumFmt
                End If
                If .HasMajorGridlines Then .MajorGridlines.Format.Line.Visible = msoFalse
            End With

            With axVal
                .HasTitle = False
                .Format.Line.Visible = msoFalse
                .MajorTickMark = xlNone: .MinorTickMark = xlNone
                With .TickLabels.Font
                    .Name = fontName: .Size = 13.1: .Color = GRAY_LABEL: .Bold = False
                End With
                If yAxisNumFmt <> "" Then
                    .TickLabels.NumberFormat = yAxisNumFmt
                ElseIf InStr(.TickLabels.NumberFormat, "%") > 0 Then
                    .TickLabels.NumberFormat = "0%"
                End If
                If yAxisLines > 0 Then
                    .MajorUnit = (.MaximumScale - .MinimumScale) / yAxisLines
                End If
                If Not .MajorGridlines Is Nothing Then
                    With .MajorGridlines.Format.Line
                        .Visible = msoTrue: .ForeColor.RGB = GRAY_GRID
                        .Weight = 0.5: .DashStyle = msoLineSolid
                    End With
                End If
            End With

            ' Style markers: none for plain Line, small filled circle for Line+Markers / Scatter
            If chartType = xlLine Or chartType = xlLineMarkers Or _
               chartType = xlXYScatterLines Then
                Dim sLine As Series
                For Each sLine In cht.SeriesCollection
                    If chartType = xlLine Then
                        sLine.MarkerStyle = xlMarkerStyleNone
                    Else
                        Dim mClr As Long
                        mClr = sLine.Format.Line.ForeColor.RGB
                        sLine.MarkerStyle = xlMarkerStyleCircle
                        sLine.MarkerSize = 3
                        sLine.MarkerForegroundColor = mClr
                        sLine.MarkerBackgroundColor = mClr
                    End If
                Next sLine
            End If

    End Select

    On Error GoTo 0
End Sub

' ============================================================
' COLOR SCHEMES (8 colours each)
' ============================================================
Private Sub ApplyColorScheme(cht As Chart, scheme As String, chartType As Long, lineWeight As Double)

    Dim palettes As Variant

    Select Case scheme

        Case "Canva"
            If chartType = xlDoughnut Or chartType = xlPie Then
                palettes = Array( _
                    RGB(111, 62, 59), RGB(129, 126, 111), RGB(201, 195, 183), _
                    RGB(120, 132, 132), RGB(77, 78, 80), RGB(129, 126, 111), _
                    RGB(201, 195, 183), RGB(120, 132, 132))
            Else
                palettes = Array( _
                    RGB(129, 126, 111), RGB(201, 195, 183), RGB(30, 76, 96), _
                    RGB(138, 125, 73), RGB(111, 62, 59), RGB(120, 132, 132), _
                    RGB(77, 78, 80), RGB(201, 195, 183))
            End If

        Case "Office"
            palettes = Array( _
                RGB(68, 114, 196), RGB(237, 125, 49), RGB(165, 165, 165), _
                RGB(255, 192, 0), RGB(91, 155, 213), RGB(112, 173, 71), _
                RGB(38, 68, 120), RGB(158, 72, 14))

        Case "Vivid"
            palettes = Array( _
                RGB(255, 87, 51), RGB(51, 181, 229), RGB(255, 195, 0), _
                RGB(76, 187, 23), RGB(142, 68, 173), RGB(26, 188, 156), _
                RGB(231, 76, 60), RGB(52, 152, 219))

        Case "Pastel"
            palettes = Array( _
                RGB(174, 214, 241), RGB(250, 215, 160), RGB(171, 235, 198), _
                RGB(245, 183, 177), RGB(215, 189, 226), RGB(249, 231, 159), _
                RGB(209, 236, 241), RGB(253, 212, 230))

        Case "Greyscale"
            palettes = Array( _
                RGB(40, 40, 40), RGB(85, 85, 85), RGB(130, 130, 130), _
                RGB(170, 170, 170), RGB(200, 200, 200), RGB(60, 60, 60), _
                RGB(110, 110, 110), RGB(150, 150, 150))

        Case "Dark"
            palettes = Array( _
                RGB(0, 173, 181), RGB(255, 170, 51), RGB(220, 80, 80), _
                RGB(100, 200, 120), RGB(160, 100, 220), RGB(240, 110, 160), _
                RGB(50, 150, 200), RGB(200, 200, 60))

        Case Else
            Exit Sub

    End Select

    Dim i As Integer
    Dim s As Series
    i = 0

    If chartType = xlPie Or chartType = xlDoughnut Then
        ' Color each slice (point) individually
        Dim s1 As Series
        Set s1 = cht.SeriesCollection(1)
        Dim pt As Point
        Dim ptIdx As Integer
        ptIdx = 0
        For Each pt In s1.Points
            Dim pClr As Long
            pClr = palettes(ptIdx Mod (UBound(palettes) + 1))
            pt.Format.Fill.ForeColor.RGB = pClr
            pt.Format.Line.ForeColor.RGB = pClr
            ptIdx = ptIdx + 1
        Next pt
    Else
        For Each s In cht.SeriesCollection
            Dim clr As Long
            clr = palettes(i Mod (UBound(palettes) + 1))
            s.Format.Fill.ForeColor.RGB = clr
            s.Format.Line.ForeColor.RGB = clr
            If chartType = xlLine Or chartType = xlLineMarkers Or _
               chartType = xlXYScatterLines Then
                s.Format.Line.Weight = lineWeight
                If chartType = xlLine Then
                    s.MarkerStyle = xlMarkerStyleNone
                ElseIf chartType = xlLineMarkers Or chartType = xlXYScatterLines Then
                    s.MarkerStyle = xlMarkerStyleCircle
                    s.MarkerSize = 3
                    s.MarkerForegroundColor = clr
                    s.MarkerBackgroundColor = clr
                End If
            End If
            i = i + 1
        Next s
    End If

End Sub

' ============================================================
Private Sub FormatAxes(cht As Chart, fontName As String, Optional yAxisNumFmt As String = "", Optional yAxisLines As Integer = 0, Optional xAxisNumFmt As String = "")
    On Error Resume Next
    Dim axCat As Axis, axVal As Axis
    Set axCat = cht.Axes(xlCategory)
    Set axVal = cht.Axes(xlValue)

    With axCat
        .HasTitle = False
        .TickLabels.Font.Name = fontName
        .TickLabels.Font.Size = 13.1
        .TickLabels.Font.Color = RGB(51, 51, 51)
        .AxisBetweenCategories = True
        If xAxisNumFmt <> "" Then
            On Error Resume Next
            .AxisType = xlTimeScale
            On Error GoTo 0
            .TickLabels.NumberFormat = xAxisNumFmt
        End If
        If .HasMajorGridlines Then .MajorGridlines.Format.Line.Visible = msoFalse
    End With

    With axVal
        .HasTitle = False
        .TickLabels.Font.Name = fontName
        .TickLabels.Font.Size = 13.1
        .TickLabels.Font.Color = RGB(51, 51, 51)
        If yAxisNumFmt <> "" Then
            .TickLabels.NumberFormat = yAxisNumFmt
        ElseIf InStr(.TickLabels.NumberFormat, "%") > 0 Then
            .TickLabels.NumberFormat = "0%"
        End If
        If yAxisLines > 0 Then
            .MajorUnit = (.MaximumScale - .MinimumScale) / yAxisLines
        End If
        With .MajorGridlines.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(200, 200, 200)
            .DashStyle = msoLineDash
        End With
    End With
    On Error GoTo 0
End Sub

' ============================================================
Public Function ChartTypeFromName(n As String) As Long
    Select Case n
        Case "Bar (Horizontal)":    ChartTypeFromName = xlBarClustered
        Case "Column (Vertical)":   ChartTypeFromName = xlColumnClustered
        Case "Line":                ChartTypeFromName = xlLine
        Case "Line with Markers":   ChartTypeFromName = xlLineMarkers
        Case "Pie":                 ChartTypeFromName = xlPie
        Case "Doughnut":            ChartTypeFromName = xlDoughnut
        Case "Area":                ChartTypeFromName = xlArea
        Case "Scatter":             ChartTypeFromName = xlXYScatter
        Case "Scatter with Lines":  ChartTypeFromName = xlXYScatterLines
        Case "Radar":               ChartTypeFromName = xlRadar
        Case "Stacked Bar":         ChartTypeFromName = xlBarStacked
        Case "Stacked Column":      ChartTypeFromName = xlColumnStacked
        Case Else:                  ChartTypeFromName = xlColumnClustered
    End Select
End Function

' ============================================================
' Color helpers
' ============================================================
Private Function ParseColor(s As String) As Long
    On Error GoTo Fail
    s = Trim(s)
    If Left(s, 1) = "#" And Len(s) = 7 Then
        Dim r As Integer, g As Integer, b As Integer
        r = CInt("&H" & Mid(s, 2, 2))
        g = CInt("&H" & Mid(s, 4, 2))
        b = CInt("&H" & Mid(s, 6, 2))
        ParseColor = RGB(r, g, b)
        Exit Function
    End If
    If InStr(s, ",") > 0 Then
        Dim parts() As String
        parts = Split(s, ",")
        If UBound(parts) = 2 Then
            ParseColor = RGB(CInt(Trim(parts(0))), CInt(Trim(parts(1))), CInt(Trim(parts(2))))
            Exit Function
        End If
    End If
Fail:
    ParseColor = -1
End Function

Private Function ColorToHex(clr As Long) As String
    Dim r As Long, g As Long, b As Long
    b = (clr \ 65536) And 255
    g = (clr \ 256) And 255
    r = clr And 255
    ColorToHex = "#" & Right("0" & Hex(r), 2) & Right("0" & Hex(g), 2) & Right("0" & Hex(b), 2)
End Function
