Private Sub AddBtn_Click()
    
    'Setting variables
    pageVal = UserForm.TextBox1.Text
    propVal = UserForm.TextBox2.Text
    pCount = Application.WorksheetFunction.CountA(Range("A1").EntireColumn)
    projPrefix = Range("E" & pCount + 2).Value

    If pageVal = "" Then
    Else
        'Adding new row
        Rows(pCount + 1 & ":" & pCount + 1).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
        'Adding data to new row depending on state of checkbox
        If UserForm.CheckBox1.Value = False Then
            Range("A" & pCount + 1) = pageVal
            Range("C" & pCount + 1) = pageVal
            If Len(propVal) = 3 Then
                Range("B" & pCount + 1) = projPrefix + propVal
            ElseIf Len(propVal) = 2 Then
                Range("B" & pCount + 1) = projPrefix + "0" + propVal
            Else
                Range("B" & pCount + 1) = projPrefix + "00" + propVal
            End If
            UserForm.TextBox1.Text = ""
            UserForm.TextBox2.Text = ""
        Else
            Range("A" & pCount + 1) = pageVal
            Range("C" & pCount + 1) = pageVal
            Range("B" & pCount + 1).Select

            For x = 1 To pCount
                If Range("C" & x).Value = (pageVal - 1) Then
                    Range("B" & pCount + 1) = Range("B" & x).Value
                End If
            Next x

        End If
    End If

End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub DelBtn_Click()

    'Deleting last row
    pCount = Application.WorksheetFunction.CountA(Range("B1").EntireColumn)
    Rows(pCount & ":" & pCount).Select
    Selection.Delete

End Sub

Private Sub Label1_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Click()

End Sub