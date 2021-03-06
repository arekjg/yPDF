Sub macro2(projectNum As String, filePath As String)
'
' macro2
'

    'Deleting all existing data from sheet
    Cells.Select
    Selection.ClearContents

    'Copying data from csv file
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & filePath, Destination:=Range("$A$1"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1250
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

    'Setting variables
    char = projectNum & "-"
    projPage = projectNum & "/"
    
    'Placing variables in cells
    Range("AM1001") = projPage
    Range("AN1001") = char
       
    'Formulas - finding columns
    Range("R1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR((ISNUMBER(FIND(R1001C39,RC[-17]))),(ISNUMBER(FIND(R1001C40,RC[-17])))),1,0)"
    Range("R1").Select
    Selection.AutoFill Destination:=Range("R1:AF1"), Type:=xlFillDefault
    Range("R1:AF1").Select
    Range("AH1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-16]=1,RC[-15]=1,RC[-14]=1,RC[-13]=1,RC[-12]=1,RC[-11]=1,RC[-10]=1,RC[-9]=1,RC[-8]=1,RC[-7]=1,RC[-6]=1),1,""x"")"
    
    'Fomulas - counting columns
    Range("AI1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-17]=1,1,IF(RC[-16]=1,2,IF(RC[-15]=1,3,IF(RC[-14]=1,4,IF(RC[-13]=1,5,IF(RC[-12]=1,6,IF(RC[-11]=1,7,IF(RC[-10]=1,8,IF(RC[-9]=1,9,IF(RC[-8]=1,10,IF(RC[-7]=1,11,IF(RC[-6]=1,12,IF(RC[-5]=1,13,IF(RC[-4]=1,14,IF(RC[-3]=1,15,0)))))))))))))))"
    Range("AJ1").Select
    ActiveCell.FormulaR1C1 = _
        "=16-(IF(RC[-4]=1,1,IF(RC[-5]=1,2,IF(RC[-6]=1,3,IF(RC[-7]=1,4,IF(RC[-8]=1,5,IF(RC[-9]=1,6,IF(RC[-10]=1,7,IF(RC[-11]=1,8,IF(RC[-12]=1,9,IF(RC[-13]=1,10,IF(RC[-14]=1,11,IF(RC[-15]=1,12,IF(RC[-16]=1,13,IF(RC[-17]=1,14,IF(RC[-18]=1,15,0))))))))))))))))"
    Range("AI1:AJ1").Select
    Selection.AutoFill Destination:=Range("AI1:AJ1000"), Type:=xlFillDefault
    Range("AI1:AJ1000").Select
    Range("AI1").Select
        
    'Deleting rows (with bool=0)
    Range("R1:AH1").Select
    Selection.AutoFill Destination:=Range("R1:AH1000"), Type:=xlFillDefault
    Range("R1:AH1000").Select
    Columns("AH:AH").Select
    Selection.SpecialCells(xlCellTypeFormulas, 2).Select
    Selection.EntireRow.Delete
    Range("AH1").Select
    
    'Copying important columns
    colStart = Range("AI1").Value
    colEnd = Range("AJ1").Value
    Range(Cells(1, colStart), Cells(1, colEnd)).EntireColumn.Select
    Selection.Copy
    Range("AK1").Select
    ActiveSheet.Paste
    Columns("A:AJ").Select
    Range("AJ1").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    
    'Managing page numbers
    pCount = Application.WorksheetFunction.CountA(Range("A1").EntireColumn)
    Range("C1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(RC[-2],R" & pCount + 1 & "C3,""""),""/00"",""""),""/0"","""")"
    Range("C1").Select
    Selection.AutoFill Destination:=Range("C1:C" & pCount), Type:=xlFillDefault
    
    '---------------------------

    'Managing page numbers (continued)
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    'Copying page numbers to column A
    Range("C1:C" & pCount).Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Deleting useless column
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    'Moving project number
    Range("C" & pCount + 1).Select
    Selection.Copy
    Range("E" & pCount + 2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C" & pCount + 1).Select
    Selection.Clear
    
    'Sorting
    Columns("A:C").Select
    Range("A1").Activate
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields. _
        Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A1:C1001")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Inserting calculation functions
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-2]+RC[2],RC[1])"
    Range("C1").Select
    Selection.AutoFill Destination:=Range("C1:C" & pCount), Type:=xlFillDefault
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]="""",0,RC[-1]-RC[-4])"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]="""",R[-1]C,RC[-1]-RC[-4])"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E" & pCount), Type:=xlFillDefault

    'Adjusting columns, hiding column E
    Range("B1:B1").Columns.AutoFit
    Columns("A:A").Select
    Selection.ColumnWidth = 17
    Columns("E:E").Select
    Selection.EntireColumn.Hidden = True

    'Filling column D with color, drawing cell borders
    Range("D1:D" & pCount).Select
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
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("D1").Select

    '----------------------------

    'BUTTONS
    'Adding "open form" button
    Dim formBtn As Button
    Application.ScreenUpdating = False
    pCount = Application.WorksheetFunction.CountA(Range("A1").EntireColumn)
    Set formBtn = ActiveSheet.Buttons.Add(300, pCount * 15 + 40, 100, 25)
    formBtn.Characters.Text = "open form"
    formBtn.OnAction = "Form_Click"
    
    'Adding "finish" button
    Dim sqBtn As Button
    Application.ScreenUpdating = False
    pCount = Application.WorksheetFunction.CountA(Range("A1").EntireColumn)
    Set formBtn = ActiveSheet.Buttons.Add(300, pCount * 15 + 66, 100, 25)
    formBtn.Characters.Text = "finish"
    formBtn.OnAction = "Finish_Click"

    Range("C" & pCount + 1).Select

    
End Sub