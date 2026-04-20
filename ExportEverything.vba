Private Const COL_SHEET_NAME As Integer = 1
Private Const COL_ADDRESS As Integer = 2
Private Const COL_VALUE As Integer = 3
Private Const COL_FORMULA As Integer = 4
Private Const COL_ROW As Integer = 5
Private Const COL_COL As Integer = 6
Private Const COL_ROW_HEIGHT As Integer = 7
Private Const COL_COL_WIDTH As Integer = 8
Private Const COL_WIDTH_POINTS As Integer = 9
Private Const COL_IS_MERGED As Integer = 10
Private Const COL_MERGE_AREA As Integer = 11
Private Const COL_ROW_HIDDEN As Integer = 12
Private Const COL_H_ALIGN As Integer = 13
Private Const COL_V_ALIGN As Integer = 14
Private Const COL_WRAP_TEXT As Integer = 15
Private Const COL_SHRINK_TO_FIT As Integer = 16
Private Const COL_INDENT As Integer = 17
Private Const COL_NUM_FORMAT As Integer = 18
Private Const COL_FONT_NAME As Integer = 19
Private Const COL_FONT_SIZE As Integer = 20
Private Const COL_FONT_BOLD As Integer = 21
Private Const COL_FONT_ITALIC As Integer = 22
Private Const COL_FONT_UNDERLINE As Integer = 23
Private Const COL_FONT_COLOR As Integer = 24
Private Const COL_FONT_THEME As Integer = 25
Private Const COL_FONT_TINT As Integer = 26
Private Const COL_FILL_COLOR As Integer = 27
Private Const COL_FILL_PATTERN As Integer = 28
Private Const COL_FILL_THEME As Integer = 29
Private Const COL_FILL_TINT As Integer = 30
Private Const COL_BORDER_LEFT_STYLE As Integer = 31
Private Const COL_BORDER_LEFT_WEIGHT As Integer = 32
Private Const COL_BORDER_LEFT_COLOR As Integer = 33
Private Const COL_BORDER_TOP_STYLE As Integer = 34
Private Const COL_BORDER_TOP_WEIGHT As Integer = 35
Private Const COL_BORDER_TOP_COLOR As Integer = 36
Private Const COL_BORDER_RIGHT_STYLE As Integer = 37
Private Const COL_BORDER_RIGHT_WEIGHT As Integer = 38
Private Const COL_BORDER_RIGHT_COLOR As Integer = 39
Private Const COL_BORDER_BOTTOM_STYLE As Integer = 40
Private Const COL_BORDER_BOTTOM_WEIGHT As Integer = 41
Private Const COL_BORDER_BOTTOM_COLOR As Integer = 42
Private Const COL_DIAG_DOWN_STYLE As Integer = 43
Private Const COL_DIAG_DOWN_COLOR As Integer = 44
Private Const COL_DIAG_UP_STYLE As Integer = 45
Private Const COL_DIAG_UP_COLOR As Integer = 46
Private Const COL_LOCKED As Integer = 47
Private Const COL_FORMULA_HIDDEN As Integer = 48
Private Const COL_COMMENT As Integer = 49
Private Const COL_HYPERLINK As Integer = 50
Private Const COL_COND_FORMAT_COUNT As Integer = 51
Private Const COL_HAS_VALIDATION As Integer = 52
Private Const COL_VALIDATION_RULE As Integer = 53
Private Const COL_PREFIX_CHAR As Integer = 54
Private Const COL_ORIENTATION As Integer = 55
Private Const COL_STRIKETHROUGH As Integer = 56
Private Const COL_SUPERSCRIPT As Integer = 57
Private Const COL_SUBSCRIPT As Integer = 58
Private Const COL_READING_ORDER As Integer = 59
Private Const COL_ADD_INDENT As Integer = 60
Private Const TOTAL_COLUMNS As Integer = 60

Sub ExportEverything()
    Dim ws As Worksheet, out As Worksheet
    Dim c As Range
    Dim data() As Variant
    Dim rowOut As Long
    Dim usedRng As Range
    Dim targetRow As Long
    Dim cellCount As Long
    
    On Error GoTo ErrorHandler

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

    Set out = Worksheets.Add(Before:=Worksheets(1))
    out.Name = "Web_Audit_" & Format(Now, "hhmmss")

    Dim Headers As Variant
    Headers = GetHeaders()
    out.Range("A1").Resize(1, TOTAL_COLUMNS).Value = Headers
    targetRow = 2

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> out.Name Then
            Application.StatusBar = "Processing Sheet: " & ws.Name
            
            Set usedRng = ws.UsedRange
            
            If Not usedRng Is Nothing Then
                cellCount = usedRng.Cells.Count
                ReDim data(1 To cellCount, 1 To TOTAL_COLUMNS)
                rowOut = 0

                Dim cell As Range
                For Each cell In usedRng
                    rowOut = rowOut + 1
                    ExportCellData cell, data, rowOut, ws
                Next cell

                If rowOut > 0 Then
                    out.Cells(targetRow, 1).Resize(rowOut, TOTAL_COLUMNS).Value = data
                    targetRow = targetRow + rowOut
                End If
                
                Erase data
            End If
        End If
    Next ws

    out.Rows(1).Font.Bold = True
    out.Columns.AutoFit

    Application.StatusBar = False
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    
    MsgBox "Export Complete!", vbInformation
    Exit Sub
    
ErrorHandler:
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .StatusBar = False
    End With
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Function GetHeaders() As Variant
    GetHeaders = Array( _
        "SheetName", "Address", "Value", "Formula", "Row", "Col", _
        "RowHeight", "ColWidth", "WidthPoints", "IsMerged", "MergeArea", "RowHidden", _
        "H-Align", "V-Align", "WrapText", "ShrinkToFit", "Indent", "NumFormat", _
        "FontName", "FontSize", "Bold", "Italic", "Underline", "FontColor", "FontTheme", "FontTint", _
        "FillColor", "FillPattern", "FillTheme", "FillTint", _
        "B-LeftStyle", "B-LeftWeight", "B-LeftColor", "B-TopStyle", "B-TopWeight", "B-TopColor", _
        "B-RightStyle", "B-RightWeight", "B-RightColor", "B-BottomStyle", "B-BottomWeight", "B-BottomColor", _
        "DiagDownStyle", "DiagDownColor", "DiagUpStyle", "DiagUpColor", _
        "Locked", "FormulaHidden", "Comment", "Hyperlink", "CondFormatCount", "HasValidation", "ValidationRule", "PrefixChar", _
        "Orientation", "Strikethrough", "Superscript", "Subscript", "ReadingOrder", "AddIndent")
End Function

Sub ExportCellData(cell As Range, data() As Variant, rowIdx As Long, ws As Worksheet)
    data(rowIdx, COL_SHEET_NAME) = ws.Name
    data(rowIdx, COL_ADDRESS) = cell.Address(False, False)
    data(rowIdx, COL_VALUE) = cell.Value
    data(rowIdx, COL_FORMULA) = IIf(cell.HasFormula, cell.Formula, "")
    data(rowIdx, COL_ROW) = cell.Row
    data(rowIdx, COL_COL) = cell.Column

    data(rowIdx, COL_ROW_HEIGHT) = cell.RowHeight
    data(rowIdx, COL_COL_WIDTH) = cell.ColumnWidth
    data(rowIdx, COL_WIDTH_POINTS) = cell.Width
    data(rowIdx, COL_IS_MERGED) = cell.MergeCells
    data(rowIdx, COL_MERGE_AREA) = IIf(cell.MergeCells, cell.MergeArea.Address, "")
    data(rowIdx, COL_ROW_HIDDEN) = cell.EntireRow.Hidden

    data(rowIdx, COL_H_ALIGN) = FormatHAlign(cell.HorizontalAlignment)
    data(rowIdx, COL_V_ALIGN) = FormatVAlign(cell.VerticalAlignment)
    data(rowIdx, COL_WRAP_TEXT) = cell.WrapText
    data(rowIdx, COL_SHRINK_TO_FIT) = cell.ShrinkToFit
    data(rowIdx, COL_INDENT) = cell.IndentLevel

    data(rowIdx, COL_NUM_FORMAT) = cell.NumberFormat

    ExportFontData cell, data, rowIdx

    ExportFillData cell, data, rowIdx

    ExportBorderData cell, data, rowIdx

    data(rowIdx, COL_LOCKED) = cell.Locked
    data(rowIdx, COL_FORMULA_HIDDEN) = cell.FormulaHidden
    data(rowIdx, COL_COMMENT) = GetCellComment(cell)
    data(rowIdx, COL_HYPERLINK) = GetCellHyperlink(cell)
    data(rowIdx, COL_COND_FORMAT_COUNT) = cell.FormatConditions.Count

    ExportValidationData cell, data, rowIdx

    data(rowIdx, COL_PREFIX_CHAR) = cell.PrefixCharacter
    data(rowIdx, COL_ORIENTATION) = cell.Orientation
    data(rowIdx, COL_READING_ORDER) = cell.ReadingOrder
    data(rowIdx, COL_ADD_INDENT) = cell.AddIndent
End Sub

Sub ExportFontData(cell As Range, data() As Variant, rowIdx As Long)
    With cell.Font
        data(rowIdx, COL_FONT_NAME) = .Name
        data(rowIdx, COL_FONT_SIZE) = .Size
        data(rowIdx, COL_FONT_BOLD) = .Bold
        data(rowIdx, COL_FONT_ITALIC) = .Italic
        data(rowIdx, COL_FONT_UNDERLINE) = .Underline
        data(rowIdx, COL_FONT_COLOR) = FormatColor(.Color)
        data(rowIdx, COL_STRIKETHROUGH) = .Strikethrough
        data(rowIdx, COL_SUPERSCRIPT) = .Superscript
        data(rowIdx, COL_SUBSCRIPT) = .Subscript
    End With

    data(rowIdx, COL_FONT_THEME) = GetSafeThemeColor(cell.Font)
    data(rowIdx, COL_FONT_TINT) = GetSafeTintAndShade(cell.Font)
End Sub

Sub ExportFillData(cell As Range, data() As Variant, rowIdx As Long)
    With cell.Interior
        data(rowIdx, COL_FILL_COLOR) = FormatColor(.Color)
        data(rowIdx, COL_FILL_PATTERN) = .Pattern
    End With
    
    data(rowIdx, COL_FILL_THEME) = GetSafeThemeColor(cell.Interior)
    data(rowIdx, COL_FILL_TINT) = GetSafeTintAndShade(cell.Interior)
End Sub

Sub ExportBorderData(cell As Range, data() As Variant, rowIdx As Long)
    With cell.Borders(xlEdgeLeft)
        data(rowIdx, COL_BORDER_LEFT_STYLE) = .LineStyle
        data(rowIdx, COL_BORDER_LEFT_WEIGHT) = .Weight
        data(rowIdx, COL_BORDER_LEFT_COLOR) = FormatColor(.Color)
    End With

    With cell.Borders(xlEdgeTop)
        data(rowIdx, COL_BORDER_TOP_STYLE) = .LineStyle
        data(rowIdx, COL_BORDER_TOP_WEIGHT) = .Weight
        data(rowIdx, COL_BORDER_TOP_COLOR) = FormatColor(.Color)
    End With

    With cell.Borders(xlEdgeRight)
        data(rowIdx, COL_BORDER_RIGHT_STYLE) = .LineStyle
        data(rowIdx, COL_BORDER_RIGHT_WEIGHT) = .Weight
        data(rowIdx, COL_BORDER_RIGHT_COLOR) = FormatColor(.Color)
    End With

    With cell.Borders(xlEdgeBottom)
        data(rowIdx, COL_BORDER_BOTTOM_STYLE) = .LineStyle
        data(rowIdx, COL_BORDER_BOTTOM_WEIGHT) = .Weight
        data(rowIdx, COL_BORDER_BOTTOM_COLOR) = FormatColor(.Color)
    End With

    With cell.Borders(xlDiagonalDown)
        data(rowIdx, COL_DIAG_DOWN_STYLE) = .LineStyle
        data(rowIdx, COL_DIAG_DOWN_COLOR) = FormatColor(.Color)
    End With
    
    With cell.Borders(xlDiagonalUp)
        data(rowIdx, COL_DIAG_UP_STYLE) = .LineStyle
        data(rowIdx, COL_DIAG_UP_COLOR) = FormatColor(.Color)
    End With
End Sub

Sub ExportValidationData(cell As Range, data() As Variant, rowIdx As Long)
    Dim hasVal As Boolean
    Dim valRule As String
    
    hasVal = False
    valRule = ""
    
    On Error Resume Next
    hasVal = Not cell.Validation Is Nothing
    If hasVal Then valRule = cell.Validation.Formula1
    If Err.Number <> 0 Then valRule = "Error reading rule": Err.Clear
    On Error GoTo 0
    
    data(rowIdx, COL_HAS_VALIDATION) = hasVal
    data(rowIdx, COL_VALIDATION_RULE) = valRule
End Sub

Function GetCellComment(cell As Range) As String
    If Not cell.Comment Is Nothing Then
        GetCellComment = cell.Comment.Text
    Else
        GetCellComment = ""
    End If
End Function

Function GetCellHyperlink(cell As Range) As String
    If cell.Hyperlinks.Count > 0 Then
        GetCellHyperlink = cell.Hyperlinks(1).Address
    Else
        GetCellHyperlink = ""
    End If
End Function

Function GetSafeThemeColor(obj As Object) As Variant
    On Error Resume Next
    GetSafeThemeColor = obj.ThemeColor
    If Err.Number <> 0 Then GetSafeThemeColor = "N/A": Err.Clear
    On Error GoTo 0
End Function

Function GetSafeTintAndShade(obj As Object) As Variant
    On Error Resume Next
    GetSafeTintAndShade = obj.TintAndShade
    If Err.Number <> 0 Then GetSafeTintAndShade = "N/A": Err.Clear
    On Error GoTo 0
End Function

Function FormatColor(ByVal decColor As Variant) As String
    If IsEmpty(decColor) Or IsNull(decColor) Then
        FormatColor = "N/A"
        Exit Function
    End If

    If Not IsNumericValue(decColor) Then
        FormatColor = "N/A"
        Exit Function
    End If

    If CLng(decColor) < 0 Then
        FormatColor = CStr(CLng(decColor)) & " | Automatic"
        Exit Function
    End If

    Dim r As Long, g As Long, b As Long
    Dim decVal As Long
    decVal = CLng(decColor)
    
    r = decVal Mod 256
    g = (decVal \ 256) Mod 256
    b = (decVal \ 65536) Mod 256
    
    FormatColor = CStr(decVal) & " | #" & _
                  Right("0" & Hex(r), 2) & _
                  Right("0" & Hex(g), 2) & _
                  Right("0" & Hex(b), 2)
End Function

Function IsNumericValue(val As Variant) As Boolean
    If IsEmpty(val) Or IsNull(val) Then
        IsNumericValue = False
        Exit Function
    End If
    
    On Error Resume Next
    Dim test As Double
    test = CDbl(val)
    IsNumericValue = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Function FormatHAlign(ByVal val As Variant) As String
    Select Case val
        Case 1: FormatHAlign = "1 | General"
        Case -4131: FormatHAlign = "-4131 | Left"
        Case -4108: FormatHAlign = "-4108 | Center"
        Case -4152: FormatHAlign = "-4152 | Right"
        Case 5: FormatHAlign = "5 | Fill"
        Case 7: FormatHAlign = "7 | CenterAcrossSelection"
        Case -4130: FormatHAlign = "-4130 | Justify"
        Case -4117: FormatHAlign = "-4117 | Distributed"
        Case Else: FormatHAlign = val & " | Unknown"
    End Select
End Function

Function FormatVAlign(ByVal val As Variant) As String
    Select Case val
        Case -4160: FormatVAlign = "-4160 | Top"
        Case -4108: FormatVAlign = "-4108 | Center"
        Case -4107: FormatVAlign = "-4107 | Bottom"
        Case -4130: FormatVAlign = "-4130 | Justify"
        Case -4117: FormatVAlign = "-4117 | Distributed"
        Case Else: FormatVAlign = val & " | Unknown"
    End Select
End Function
