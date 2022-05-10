Sub Finish_Click()
    
    pCount = Application.WorksheetFunction.CountA(Range("A1").EntireColumn)

    'Copying page numbers to column A
    Range("C1:C" & pCount).Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'Converting to numbers
    Range("A1:A" & pCount).Select
    With Selection
        .NumberFormat = "General"
        .Value = .Value
    End With
    
    'Deleting useless columns
    Columns("C:I").Select
    Selection.Delete Shift:=xlToLeft

    'Sorting alphabetically by column B
    Columns("A:B").Select
    Range("B1").Activate
    ActiveWorkbook.Worksheets("Sheet1").sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").sort.SortFields.Add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").sort
        .SetRange Range("A1:B" & pCount)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Saving and closing
    Application.DisplayAlerts = False
    Workbooks("temp.xlsm").SaveAs FileName:= Workbooks("temp.xlsm").Path & "\temp_b.xlsm"


End Sub