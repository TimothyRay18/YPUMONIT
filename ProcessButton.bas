Attribute VB_Name = "ProcessButton"
Function getMaxRow(col As Integer) As Double
    getMaxRow = ActiveSheet.Cells(Rows.Count, col).End(xlUp).row
End Function

Function getMaxCol(row As Integer) As Double
    getMaxCol = ActiveSheet.Cells(row, Columns.Count).End(xlToLeft).Column
End Function

Function GetFilenameFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

Function findCellInColumn(row As Integer, str As String) As Double
    Dim i As Double
    i = 1
    Dim m As Double
    m = getMaxCol(row)
    While LCase(ActiveSheet.Cells(row, i).Value) <> LCase(str) And i <= m
        i = i + 1
    Wend
    findCellInColumn = i
End Function

Sub addBorder(start As String, last As String)
    Range(start & ":" & last).Select
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Sub Process()
    OpenAndClean
    ProcessTemplate
End Sub

Sub OpenAndClean()
    Dim macro_file As String
    macro_file = ActiveWorkbook.Name

    Dim dir1 As String
    dir1 = Range("B1").Value
    Workbooks.OpenText Filename:= _
        dir1, Origin:=xlWindows, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 2), Array(4, 2), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 2), _
        Array(10, 1), Array(11, 1), Array(12, 2), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), _
        Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array( _
        29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), _
        Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array( _
        42, 1), Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 1), _
        Array(49, 1), Array(50, 1), Array(51, 1), Array(52, 1), Array(53, 1), Array(54, 1), Array( _
        55, 1), Array(56, 1), Array(57, 1), Array(58, 1), Array(59, 1), Array(60, 1), Array(61, 1), _
        Array(62, 1), Array(63, 1), Array(64, 1), Array(65, 1), Array(66, 1), Array(67, 1), Array( _
        68, 1), Array(69, 1), Array(70, 1), Array(71, 1), Array(72, 1), Array(73, 1), Array(74, 1), _
        Array(75, 1), Array(76, 1), Array(77, 1), Array(78, 1), Array(79, 1), Array(80, 1), Array( _
        81, 1), Array(82, 1), Array(83, 1)), TrailingMinusNumbers:=True

    Dim source_file As String
    source_file = ActiveWorkbook.Name
    
'    If IsEmpty(Range("A1").Value) = True Then
'        Exit Sub
'    End If
'    If IsEmpty(Range("A1").Value) = False Then
    Rows("1:6").Select
    Selection.Delete Shift:=xlUp
    Range("C1").Value = "Material"
    Columns("A:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("AD:AD").Select
    Selection.Delete Shift:=xlToLeft
    Columns("AF:AF").Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
'    End If
    
    Dim max_row As Double
    max_row = getMaxRow(1)
    Dim max_col As Double
    max_col = getMaxCol(1)

'    Range(Columns(1), Columns(max_col)).EntireColumn.AutoFit

    Range(Cells(1, 1), Cells(max_row, max_col)).Select
    Selection.AutoFilter
    
'    Dim mrp_col As Double
'    mrp_col = findCellInColumn(1, "MRP Type")
'
'    MsgBox mrp_col

    ActiveSheet.Range(Cells(1, 1), Cells(max_row, max_col)).AutoFilter Field:=9, Criteria1:=Array( _
        "MRP Type", "Y0", "="), Operator:=xlFilterValues
    Rows("2:" + CStr(max_row)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range(Cells(1, 1), Cells(max_row, max_col)).AutoFilter Field:=9
    Selection.AutoFilter

    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Columns("A:A").Select
    Selection.Copy
    Columns("B:B").Select
    ActiveSheet.Paste
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

    Dim wb As Workbook
    Dim opened_m0 As String
    
    Workbooks(macro_file).Activate
    Set wb = Workbooks.Open(Range("B6").Value)
    wb.Worksheets("Sheet1").Activate
    m0_file = ActiveWorkbook.Name

    Workbooks(source_file).Activate

    Columns("C:C").Select
    Selection.NumberFormat = "General"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],'[" + m0_file + "]Sheet1'!C1,1,0)"
    Range("C2").Select
    max_row = getMaxRow(1)
    Selection.AutoFill Destination:=Range("C2:C" + CStr(max_row))

    Application.DisplayAlerts = False
    Windows(m0_file).Close
    Application.DisplayAlerts = True
    
'    ActiveSheet.Range(Cells(1, 1), Cells(max_row, max_col)).AutoFilter Field:=9
    
    
'    Dim fu_col As Double
'    fu_col = findCellInColumn(1, "Follow up Material")
'
'    MsgBox fu_col
    ActiveSheet.Range("$A$1:$BY$" + CStr(max_row)).AutoFilter Field:=3, Criteria1:= _
        "<>#N/A", Operator:=xlAnd
    Range("A2:C" + CStr(max_row)).Select
'    Range
    Selection.Interior.Color = RGB(255, 185, 0)
    
    Columns("B:C").Select
    Selection.Delete Shift:=xlToLeft
    
    ActiveSheet.Range("$A$1:$BW$" + CStr(max_row)).AutoFilter Field:=9, Criteria1:="M0"
    ActiveSheet.Range("$A$1:$BW$" + CStr(max_row)).AutoFilter Field:=1, Operator:= _
        xlFilterNoFill
    Rows("2:" + CStr(max_row)).Select
    Selection.Delete Shift:=xlUp
    
    ActiveSheet.ShowAllData
    max_row = getMaxRow(1)
    ActiveSheet.Range("$A$1:$BW$" + CStr(max_row)).AutoFilter Field:=5, Criteria1:="~*"
    Range("E2:E" + CStr(max_row)).Select
    Selection.ClearContents
    
    Selection.AutoFilter
    Rows("1:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    source_name = Replace(source_file, ".xls", "")
    Dim dir As String
    dir = Workbooks(source_file).Path
    ActiveWorkbook.SaveAs Filename:= _
        dir + "\" + source_name + ".xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
End Sub

Sub ProcessTemplate()
    ThisWorkbook.Activate
    Set wb = Workbooks.Open(Range("B11").Value)
    wb.Worksheets("YPUMONIT").Activate
    Dim template_file As String
    template_file = ActiveWorkbook.Name

    Dim macro_file As String
    macro_file = ThisWorkbook.Name
    Workbooks(macro_file).Activate

    Dim source_file As String
    source_file = GetFilenameFromPath(Range("B1").Value)
    source_file = Replace(source_file, ".xls", ".xlsx")

    Workbooks(source_file).Activate
    Columns("A:F").Select
    Selection.Copy
    Workbooks(template_file).Activate
    Columns("A:F").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Dim max_row1 As Double
    Dim max_row2 As Double
    max_row1 = getMaxRow(1)
    max_row2 = getMaxRow(7)
    
    If max_row1 < max_row2 Then
        Rows(CStr(max_row1 + 1) + ":" + CStr(max_row2)).Select
        Selection.Delete Shift:=xlUp
    End If
    
'    Range(Cells(1, 30).EntireColumn, Cells(1, sim_col - 1).EntireColumn).Select
    
    Dim vsn_col As Double
    vsn_col = findCellInColumn(5, "Vendor Short Name")
    Dim fgi_col As Double
    fgi_col = findCellInColumn(5, "FG impact")
    Range(Cells(6, vsn_col), Cells(6, fgi_col)).Select
'    Range("G6:J6").Select
    Selection.AutoFill Destination:=Range(Cells(6, vsn_col), Cells(max_row1, fgi_col))
'    Range(Cells(6, vsn_col), Cells(max_row1, fgi_col)).Select
'    Range("G6:J" + CStr(max_row1)).Select
    
    Workbooks(source_file).Activate
    Range(Cells(1, findCellInColumn(5, "Text")).EntireColumn, Cells(1, findCellInColumn(5, "Rounding Value")).EntireColumn).Select
'    Columns("G:N").Select
    Selection.Copy
    Workbooks(template_file).Activate
    Range(Cells(1, findCellInColumn(5, "Text")).EntireColumn, Cells(1, findCellInColumn(5, "Rounding Value")).EntireColumn).Select
'    Columns("K:R").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Dim sc_col As Double
    sc_col = findCellInColumn(5, "Standard Cost")
    Cells(6, sc_col).Select
'    Range("S6:S6").Select
    Selection.AutoFill Destination:=Range(Cells(6, sc_col), Cells(max_row1, sc_col))
'    Selection.AutoFill Destination:=Range("S6:S" + CStr(max_row1))
'    Range("S6:S" + CStr(max_row1)).Select
    
    Workbooks(source_file).Activate
    Columns("P:Z").Select
    Selection.Copy
    Workbooks(template_file).Activate
    Columns("T:AD").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("AF6:AF6").Select
    Selection.AutoFill Destination:=Range("AF6:AF" + CStr(max_row1))
    Range("AF6:AF" + CStr(max_row1)).Select
    
    Workbooks(source_file).Activate
    Columns("AB:AC").Select
    Selection.Copy
    Workbooks(template_file).Activate
    Columns("AG:AH").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("AJ6:AL6").Select
    Selection.AutoFill Destination:=Range("AJ6:AL" + CStr(max_row1))
    Range("AJ6:AL" + CStr(max_row1)).Select
    
    Range("AI6").Select
    Range("AI6:AI" + CStr(getMaxRow(35))).Select
    Selection.ClearContents
    
    Workbooks(source_file).Activate
    Columns("AB:AC").Select
    Selection.Copy
    Workbooks(template_file).Activate
    Columns("AG:AH").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Workbooks(source_file).Activate
    Dim sim_col As Double
    sim_col = findCellInColumn(5, "Backlog for SIM")
    Range(Cells(1, 30).EntireColumn, Cells(1, sim_col - 1).EntireColumn).Select
    Selection.Copy
    Workbooks(template_file).Activate
    Range(Cells(1, 39).EntireColumn, Cells(1, (sim_col - 1 - 30 + 39)).EntireColumn).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("BS6:DA6").Select
    Selection.AutoFill Destination:=Range("BS6:DA" + CStr(max_row1))
    Range("BS6:DA" + CStr(max_row1)).Select
    
    Dim pob_col As Double
    pob_col = findCellInColumn(5, "PO Backlog")
    Range(Cells(1, pob_col).EntireColumn, Cells(1, pob_col + 15).EntireColumn).Select
    Selection.Copy
    Range(Cells(1, 118).EntireColumn, Cells(1, 118 + 15).EntireColumn).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range(Cells(6, pob_col), Cells(max_row1, pob_col + 15)).Select
    Selection.ClearContents
    
    Range("BS5").Select
    Selection.Copy
    Range("CI5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("BS5").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "SS Backlog"
    
    Dim pic As Double
    pic = findCellInColumn(5, "PIC")
    Columns(pic + 1).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(5, pic + 1).Select
    ActiveCell.FormulaR1C1 = "2nd PIC"
    Cells(6, pic).Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-110],[" + macro_file + "]PIC!C1:C2,2,0)"
    Selection.AutoFill Destination:=Range(Cells(6, pic), Cells(max_row1, pic))
    Cells(6, pic + 1).Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-111],[" + macro_file + "]PIC!C1:C3,3,0)"
    Selection.AutoFill Destination:=Range(Cells(6, pic + 1), Cells(max_row1, pic + 1))
    
    Range(Cells(1, pic).EntireColumn, Cells(1, pic + 1).EntireColumn).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets(Array("YPUMONIT", "NET DEMAND")).Select
    Sheets(Array("YPUMONIT", "NET DEMAND")).Copy
    
    max_col = getMaxCol(5)
    Range(Cells(5, 1), Cells(5, max_col)).Select
    Selection.AutoFilter
    Range(Cells(5, 1), Cells(max_row1, max_col)).Select
    ActiveWorkbook.Worksheets("YPUMONIT").Sort.SortFields.Clear
    pic = findCellInColumn(5, "PIC")
    ActiveWorkbook.Worksheets("YPUMONIT").Sort.SortFields.Add2 Key:=Range( _
        Cells(6, pic), Cells(max_row1, pic)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
        
'    Dim vsn_col As Double
    vsn_col = findCellInColumn(5, "Vendor Short Name")
    ActiveWorkbook.Worksheets("YPUMONIT").Sort.SortFields.Add2 Key:=Range( _
        Cells(6, vsn_col), Cells(max_row1, vsn_col)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    
    ActiveWorkbook.Worksheets("YPUMONIT").Sort.SortFields.Add2 Key:=Range( _
        Cells(6, 1), Cells(max_row1, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("YPUMONIT").Sort
        .SetRange Range(Cells(5, 1), Cells(max_row1, max_col))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'    Columns("CZ:DA").Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
    RemoveFormula
    Workbooks(source_file).Save
    Workbooks(source_file).Close
'    Workbooks(template_file).Close
End Sub
Sub RemoveFormula()
    Dim ws As Worksheet

    Set ws = ActiveWorkbook.Sheets("YPUMONIT")

    With ws.UsedRange
        .Copy
        .PasteSpecial Paste:=xlPasteValues, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End With
End Sub
