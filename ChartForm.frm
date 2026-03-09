VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChartForm
   Caption         =   "Chart Creator"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
' ChartForm.frm
' UserForm for the Chart Creator Add-In
'
' NOTE: Because .frx (binary) is not included here, import
' ChartFormCode.bas instead and let InitForm() build the
' controls programmatically — see README for instructions.
' ============================================================
Option Explicit

' ---- Control references (populated by InitForm) ----
Private lblRange        As MSForms.Label
Private txtRange        As MSForms.TextBox
Private btnPickRange    As MSForms.CommandButton

Private lblChartType    As MSForms.Label
Private cboChartType    As MSForms.ComboBox

Private lblTitle        As MSForms.Label
Private txtTitle        As MSForms.TextBox

Private lblPlacement    As MSForms.Label
Private cboPlacement    As MSForms.ComboBox

Private lblColor        As MSForms.Label
Private cboColor        As MSForms.ComboBox

Private chkHeaders      As MSForms.CheckBox
Private chkLegend       As MSForms.CheckBox
Private chkLabels       As MSForms.CheckBox

Private lblSize         As MSForms.Label
Private lblW            As MSForms.Label
Private txtW            As MSForms.TextBox
Private lblH            As MSForms.Label
Private txtH            As MSForms.TextBox

Private btnCreate       As MSForms.CommandButton
Private btnCancel       As MSForms.CommandButton

Private lblPreview      As MSForms.Label

' ============================================================
Private Sub UserForm_Initialize()
    Me.Caption = "Chart Creator"
    Me.Width   = 380
    Me.Height  = 460
    Me.BackColor = RGB(245, 245, 248)

    BuildControls
    PopulateDefaults
End Sub

' ============================================================
Private Sub BuildControls()

    Dim y As Integer
    y = 10

    ' ----- Data Range -----
    AddLabel Me, "Data Range:", 10, y, 100, 15
    Set txtRange = AddTextBox(Me, "txtRange", 115, y, 165, 18)
    Set btnPickRange = AddButton(Me, "btnPickRange", "...", 285, y, 25, 18)
    btnPickRange.Font.Bold = True
    y = y + 28

    ' ----- Chart Type -----
    AddLabel Me, "Chart Type:", 10, y, 100, 15
    Set cboChartType = AddCombo(Me, "cboChartType", 115, y, 175, 18)
    y = y + 28

    ' ----- Chart Title -----
    AddLabel Me, "Chart Title:", 10, y, 100, 15
    Set txtTitle = AddTextBox(Me, "txtTitle", 115, y, 175, 18)
    y = y + 28

    ' ----- Placement -----
    AddLabel Me, "Placement:", 10, y, 100, 15
    Set cboPlacement = AddCombo(Me, "cboPlacement", 115, y, 175, 18)
    y = y + 28

    ' ----- Color Scheme -----
    AddLabel Me, "Color Scheme:", 10, y, 100, 15
    Set cboColor = AddCombo(Me, "cboColor", 115, y, 175, 18)
    y = y + 28

    ' ----- Size -----
    AddLabel Me, "Size (pts):", 10, y, 100, 15
    AddLabel Me, "W:", 115, y, 15, 15
    Set txtW = AddTextBox(Me, "txtW", 132, y, 65, 18)
    AddLabel Me, "H:", 204, y, 15, 15
    Set txtH = AddTextBox(Me, "txtH", 220, y, 65, 18)
    y = y + 28

    ' ----- Checkboxes -----
    Set chkHeaders = AddCheckBox(Me, "chkHeaders", "First row/column = headers", 10, y, 240, 18)
    y = y + 22
    Set chkLegend  = AddCheckBox(Me, "chkLegend",  "Show legend",               10, y, 180, 18)
    y = y + 22
    Set chkLabels  = AddCheckBox(Me, "chkLabels",  "Show data labels",          10, y, 180, 18)
    y = y + 30

    ' ----- Preview label -----
    Set lblPreview = AddLabel(Me, "", 10, y, 340, 30)
    lblPreview.WordWrap = True
    lblPreview.ForeColor = RGB(100, 100, 100)
    lblPreview.Font.Italic = True
    y = y + 38

    ' ----- Buttons -----
    Set btnCreate = AddButton(Me, "btnCreate", "Create Chart", 80, y, 100, 24)
    btnCreate.BackColor  = RGB(68, 114, 196)
    btnCreate.ForeColor  = RGB(255, 255, 255)
    btnCreate.Font.Bold  = True

    Set btnCancel = AddButton(Me, "btnCancel", "Cancel", 195, y, 80, 24)
    btnCancel.BackColor = RGB(220, 220, 220)

    Me.Height = y + 70

End Sub

' ============================================================
Private Sub PopulateDefaults()

    ' Chart types
    Dim types As Variant
    types = Array( _
        "Column (Vertical)", "Bar (Horizontal)", "Line", "Line with Markers", _
        "Pie", "Doughnut", "Area", "Scatter", "Scatter with Lines", _
        "Stacked Column", "Stacked Bar", "Radar")
    Dim t As Variant
    For Each t In types
        cboChartType.AddItem t
    Next t
    cboChartType.ListIndex = 0

    ' Placement
    cboPlacement.AddItem "Right of data"
    cboPlacement.AddItem "Below data"
    cboPlacement.AddItem "New sheet"
    cboPlacement.ListIndex = 0

    ' Color schemes  ("Canva" first = default)
    Dim schemes As Variant
    schemes = Array("Canva", "Default", "Office", "Vivid", "Pastel", "Greyscale", "Dark")
    Dim sc As Variant
    For Each sc In schemes
        cboColor.AddItem sc
    Next sc
    cboColor.ListIndex = 0  ' Canva is default

    ' Size defaults (points; 1pt ≈ 1/72 inch)
    txtW.Text = "480"
    txtH.Text = "300"

    ' Checkboxes
    chkHeaders.Value = True
    chkLegend.Value  = True
    chkLabels.Value  = False

    ' Pre-fill range from selection
    On Error Resume Next
    txtRange.Text = Selection.Address
    On Error GoTo 0

    UpdatePreview
End Sub

' ============================================================
Private Sub btnPickRange_Click()
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox( _
        Prompt:="Select your data range:", _
        Title:="Select Range", _
        Type:=8)
    On Error GoTo 0
    If Not rng Is Nothing Then
        txtRange.Text = rng.Address(External:=True)
    End If
    UpdatePreview
End Sub

' ============================================================
Private Sub cboChartType_Change()
    UpdatePreview
End Sub

' ============================================================
Private Sub btnCreate_Click()
    If Not ValidateInputs() Then Exit Sub

    Dim rng       As Range
    Dim chartType As Long
    Dim placement As String

    On Error GoTo ErrHandler
    Set rng = Application.Range(txtRange.Text)

    chartType = ChartCreator.ChartTypeFromName(cboChartType.Text)

    Select Case cboPlacement.ListIndex
        Case 0: placement = "Right"
        Case 1: placement = "Below"
        Case 2: placement = "NewSheet"
    End Select

    ChartCreator.CreateChart _
        dataRange   := rng, _
        chartType   := chartType, _
        chartTitle  := Trim(txtTitle.Text), _
        hasHeaders  := CBool(chkHeaders.Value), _
        placement   := placement, _
        showLegend  := CBool(chkLegend.Value), _
        showLabels  := CBool(chkLabels.Value), _
        colorScheme := cboColor.Text, _
        chartW      := CDbl(txtW.Text), _
        chartH      := CDbl(txtH.Text)

    Unload Me
    Exit Sub

ErrHandler:
    MsgBox "Error creating chart: " & Err.Description, vbCritical, "Chart Creator"
End Sub

' ============================================================
Private Sub btnCancel_Click()
    Unload Me
End Sub

' ============================================================
Private Sub ValidateInputs_Change()
    UpdatePreview
End Sub

Private Function ValidateInputs() As Boolean
    ValidateInputs = False

    If Trim(txtRange.Text) = "" Then
        MsgBox "Please select a data range.", vbExclamation, "Chart Creator"
        txtRange.SetFocus
        Exit Function
    End If

    Dim rng As Range
    On Error Resume Next
    Set rng = Application.Range(txtRange.Text)
    On Error GoTo 0
    If rng Is Nothing Then
        MsgBox "The range """ & txtRange.Text & """ is not valid.", vbExclamation, "Chart Creator"
        txtRange.SetFocus
        Exit Function
    End If

    If Not IsNumeric(txtW.Text) Or CDbl(txtW.Text) < 50 Then
        MsgBox "Width must be a number >= 50.", vbExclamation, "Chart Creator"
        txtW.SetFocus
        Exit Function
    End If

    If Not IsNumeric(txtH.Text) Or CDbl(txtH.Text) < 50 Then
        MsgBox "Height must be a number >= 50.", vbExclamation, "Chart Creator"
        txtH.SetFocus
        Exit Function
    End If

    ValidateInputs = True
End Function

' ============================================================
Private Sub UpdatePreview()
    Dim msg As String

    If cboChartType.ListIndex >= 0 Then
        msg = "Will create a " & cboChartType.Text & " chart"
        If Trim(txtTitle.Text) <> "" Then
            msg = msg & " titled """ & Trim(txtTitle.Text) & """"
        End If
        msg = msg & " (" & txtW.Text & " x " & txtH.Text & " pts)"
        msg = msg & " using the """ & cboColor.Text & """ colour scheme."
    End If

    lblPreview.Caption = msg
End Sub

' ============================================================
' ---- Helpers to build controls at runtime ----
' ============================================================
Private Function AddLabel(parent As Object, caption As String, _
    x As Integer, y As Integer, w As Integer, h As Integer) As MSForms.Label
    Dim lbl As MSForms.Label
    Set lbl = parent.Controls.Add("Forms.Label.1")
    With lbl
        .Caption  = caption
        .Left     = x
        .Top      = y
        .Width    = w
        .Height   = h
        .Font.Size = 9
    End With
    Set AddLabel = lbl
End Function

Private Function AddTextBox(parent As Object, name As String, _
    x As Integer, y As Integer, w As Integer, h As Integer) As MSForms.TextBox
    Dim tb As MSForms.TextBox
    Set tb = parent.Controls.Add("Forms.TextBox.1")
    With tb
        .Name   = name
        .Left   = x
        .Top    = y
        .Width  = w
        .Height = h
        .Font.Size = 9
    End With
    Set AddTextBox = tb
End Function

Private Function AddCombo(parent As Object, name As String, _
    x As Integer, y As Integer, w As Integer, h As Integer) As MSForms.ComboBox
    Dim cb As MSForms.ComboBox
    Set cb = parent.Controls.Add("Forms.ComboBox.1")
    With cb
        .Name   = name
        .Left   = x
        .Top    = y
        .Width  = w
        .Height = h
        .Font.Size = 9
        .Style  = fmStyleDropDownList
    End With
    Set AddCombo = cb
End Function

Private Function AddButton(parent As Object, name As String, caption As String, _
    x As Integer, y As Integer, w As Integer, h As Integer) As MSForms.CommandButton
    Dim btn As MSForms.CommandButton
    Set btn = parent.Controls.Add("Forms.CommandButton.1")
    With btn
        .Name    = name
        .Caption = caption
        .Left    = x
        .Top     = y
        .Width   = w
        .Height  = h
        .Font.Size = 9
    End With
    Set AddButton = btn
End Function

Private Function AddCheckBox(parent As Object, name As String, caption As String, _
    x As Integer, y As Integer, w As Integer, h As Integer) As MSForms.CheckBox
    Dim chk As MSForms.CheckBox
    Set chk = parent.Controls.Add("Forms.CheckBox.1")
    With chk
        .Name     = name
        .Caption  = caption
        .Left     = x
        .Top      = y
        .Width    = w
        .Height   = h
        .Font.Size = 9
        .BackColor = RGB(245, 245, 248)
    End With
    Set AddCheckBox = chk
End Function
