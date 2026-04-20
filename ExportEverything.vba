Sub ExportEverything()
    Dim ws As Worksheet, out As Worksheet
    Dim c As Range
    Dim data() As Variant
    Dim rowOut As Long, colCount As Integer
    Dim usedRng As Range
    Dim targetRow As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    colCount = 60
    
    Set out = Worksheets.Add(Before:=Worksheets(1))
    out.Name = "Web_Audit_" & Format(Now, "hhmmss")
    
    Dim Headers As Variant
    Headers = Array("SheetName", "Address", "Value", "Formula", "Row", "Col", "RowHeight", "ColWidth", "WidthPoints", "IsMerged", "MergeArea", "RowHidden", _
                    "H-Align", "V-Align", "WrapText", "ShrinkToFit", "Indent", "NumFormat", _
                    "FontName", "FontSize", "Bold", "Italic", "Underline", "FontColor", "FontTheme", "FontTint", _
                    "FillColor", "FillPattern", "FillTheme", "FillTint", _
                    "B-LeftStyle", "B-LeftWeight", "B-LeftColor", "B-TopStyle", "B-TopWeight", "B-TopColor", _
                    "B-RightStyle", "B-RightWeight", "B-RightColor", "B-BottomStyle", "B-BottomWeight", "B-BottomColor", _
                    "DiagDownStyle", "DiagDownColor", "DiagUpStyle", "DiagUpColor", _
                    "Locked", "FormulaHidden", "Comment", "Hyperlink", "CondFormatCount", "HasValidation", "ValidationRule", "PrefixChar", _
                    "Orientation", "Strikethrough", "Superscript", "Subscript", "ReadingOrder", "AddIndent")
    
    out.Range("A1").Resize(1, colCount).Value = Headers
    targetRow = 2
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> out.Name Then
            Application.StatusBar = "Processing Sheet: " & ws.Name
            Set usedRng = ws.UsedRange
            
            If Not usedRng Is Nothing And WorksheetFunction.CountA(usedRng) > 0 Then
                ReDim data(1 To usedRng.Cells.Count, 1 To colCount)
                rowOut = 1
                
                For Each c In usedRng
                    rowOut = rowOut + 1
                    
                    data(rowOut - 1, 1) = ws.Name
                    data(rowOut - 1, 2) = c.Address(False, False)
                    data(rowOut - 1, 3) = c.Value
                    data(rowOut - 1, 4) = IIf(c.HasFormula, c.Formula, "")
                    data(rowOut - 1, 5) = c.Row
                    data(rowOut - 1, 6) = c.Column
                    
                    data(rowOut - 1, 7) = c.RowHeight
                    data(rowOut - 1, 8) = c.ColumnWidth
                    data(rowOut - 1, 9) = c.Width
                    data(rowOut - 1, 10) = c.MergeCells
                    data(rowOut - 1, 11) = IIf(c.MergeCells, c.MergeArea.Address, "")
                    
                    data(rowOut - 1, 12) = c.EntireRow.Hidden
                    
                    data(rowOut - 1, 13) = FormatHAlign(c.HorizontalAlignment)
                    data(rowOut - 1, 14) = FormatVAlign(c.VerticalAlignment)
                    
                    data(rowOut - 1, 15) = c.WrapText
                    data(rowOut - 1, 16) = c.ShrinkToFit
                    data(rowOut - 1, 17) = c.IndentLevel
                    data(rowOut - 1, 18) = c.NumberFormat
                    
                    With c.Font
                        data(rowOut - 1, 19) = .Name
                        data(rowOut - 1, 20) = .Size
                        data(rowOut - 1, 21) = .Bold
                        data(rowOut - 1, 22) = .Italic
                        data(rowOut - 1, 23) = .Underline
                        data(rowOut - 1, 24) = FormatColor(.Color)
                        data(rowOut - 1, 56) = .Strikethrough
                        data(rowOut - 1, 57) = .Superscript
                        data(rowOut - 1, 58) = .Subscript
                    End With
                    
                    On Error Resume Next
                    data(rowOut - 1, 25) = c.Font.ThemeColor
                    If Err.Number <> 0 Then data(rowOut - 1, 25) = "N/A": Err.Clear
                    data(rowOut - 1, 26) = c.Font.TintAndShade
                    If Err.Number <> 0 Then data(rowOut - 1, 26) = "N/A": Err.Clear
                    On Error GoTo 0
                    
                    data(rowOut - 1, 27) = FormatColor(c.Interior.Color)
                    data(rowOut - 1, 28) = c.Interior.Pattern
                    
                    On Error Resume Next
                    data(rowOut - 1, 29) = c.Interior.ThemeColor
                    If Err.Number <> 0 Then data(rowOut - 1, 29) = "N/A": Err.Clear
                    data(rowOut - 1, 30) = c.Interior.TintAndShade
                    If Err.Number <> 0 Then data(rowOut - 1, 30) = "N/A": Err.Clear
                    On Error GoTo 0
                    
                    data(rowOut - 1, 31) = c.Borders(xlEdgeLeft).LineStyle
                    data(rowOut - 1, 32) = c.Borders(xlEdgeLeft).Weight
                    data(rowOut - 1, 33) = FormatColor(c.Borders(xlEdgeLeft).Color)
                    
                    data(rowOut - 1, 34) = c.Borders(xlEdgeTop).LineStyle
                    data(rowOut - 1, 35) = c.Borders(xlEdgeTop).Weight
                    data(rowOut - 1, 36) = FormatColor(c.Borders(xlEdgeTop).Color)
                    
                    data(rowOut - 1, 37) = c.Borders(xlEdgeRight).LineStyle
                    data(rowOut - 1, 38) = c.Borders(xlEdgeRight).Weight
                    data(rowOut - 1, 39) = FormatColor(c.Borders(xlEdgeRight).Color)
                    
                    data(rowOut - 1, 40) = c.Borders(xlEdgeBottom).LineStyle
                    data(rowOut - 1, 41) = c.Borders(xlEdgeBottom).Weight
                    data(rowOut - 1, 42) = FormatColor(c.Borders(xlEdgeBottom).Color)
                    
                    data(rowOut - 1, 43) = c.Borders(xlDiagonalDown).LineStyle
                    data(rowOut - 1, 44) = FormatColor(c.Borders(xlDiagonalDown).Color)
                    data(rowOut - 1, 45) = c.Borders(xlDiagonalUp).LineStyle
                    data(rowOut - 1, 46) = FormatColor(c.Borders(xlDiagonalUp).Color)
                    
                    data(rowOut - 1, 47) = c.Locked
                    data(rowOut - 1, 48) = c.FormulaHidden
                    
                    If Not c.Comment Is Nothing Then
                        data(rowOut - 1, 49) = c.Comment.Text
                    Else
                        data(rowOut - 1, 49) = ""
                    End If
                    
                    If c.Hyperlinks.Count > 0 Then
                        data(rowOut - 1, 50) = c.Hyperlinks(1).Address
                    Else
                        data(rowOut - 1, 50) = ""
                    End If
                    
                    data(rowOut - 1, 51) = c.FormatConditions.Count
                    
                    Dim hasVal As Boolean, valRule As String
                    hasVal = False
                    valRule = ""
                    On Error Resume Next
                    hasVal = Not c.Validation Is Nothing
                    If hasVal Then valRule = c.Validation.Formula1
                    If Err.Number <> 0 Then valRule = "Error reading rule": Err.Clear
                    On Error GoTo 0
                    
                    data(rowOut - 1, 52) = hasVal
                    data(rowOut - 1, 53) = valRule
                    
                    data(rowOut - 1, 54) = c.PrefixCharacter
                    data(rowOut - 1, 55) = c.Orientation
                    data(rowOut - 1, 59) = c.ReadingOrder
                    data(rowOut - 1, 60) = c.AddIndent
                Next c
                
                out.Cells(targetRow, 1).Resize(UBound(data, 1), UBound(data, 2)).Value = data
                targetRow = targetRow + UBound(data, 1)
                Erase data
            End If
        End If
    Next ws

    out.Rows(1).Font.Bold = True
    out.Columns.AutoFit
    
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Export Complete!"
End Sub

Function FormatColor(ByVal decColor As Variant) As String
    On Error Resume Next
    If IsEmpty(decColor) Or IsNull(decColor) Or Not IsNumeric(decColor) Then
        FormatColor = "N/A"
        Exit Function
    End If
    If decColor < 0 Then
        FormatColor = decColor & " | Automatic"
        Exit Function
    End If

    Dim r As Long, g As Long, b As Long
    r = decColor Mod 256
    g = (decColor \ 256) Mod 256
    b = (decColor \ 65536) Mod 256

    FormatColor = decColor & " | #" & Right("0" & Hex(r), 2) & Right("0" & Hex(g), 2) & Right("0" & Hex(b), 2)
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
