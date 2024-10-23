Attribute VB_Name = "ErrorLogging"
Sub Clear_ErrorDashboard(Optional Category As String = "")
    Set ws = ThisWorkbook.Worksheets("Errors")
    

    topRow = 6
    bottomRow = ws.Range("B" & ws.Rows.Count).End(xlUp).row
    ' Find the clear range if a category was specified
    If Not Category = "" Then
        Set result = ws.Range("A" & topRow & ":B" & bottomRow).Find(Category, LookAt:=xlWhole)
        If result Is Nothing Then
            MsgBox "Error: Attempted to clear invalid category in Errors tab: " & Category
            End
        End If
        
        ' Currently the next header is found by looking for the next cell with text in the A column.
        ' This works because the A column is empty, and the headers are merged cells in column A.
        ' If column A ever gets filled with checkboxes or anything else, this code will break!
        topRow = result.row
        nextCatRow = result.End(xlDown).row
        ' The maximum "bottom row" is the last row in the sheet, found earlier
        If nextCatRow < bottomRow Then bottomRow = nextCatRow - 1
    End If
    
    
    i = bottomRow
    ' Loops from the bottom up to avoid issues with deleted rows
    Do While i > topRow
        Set row = ws.Rows(i)
        Set cell = ws.Range("B" & i)
        If row.OutlineLevel = 1 Then
            With ws.Range("C" & i & ":D" & i)
                If Not .HasFormula Then .ClearContents
            End With
            ' If cell.Offset(0, -1).Value = True Then cell.Offset(0, -1).Value = False
        Else
            row.Delete
        End If
        i = i - 1
    Loop

End Sub

Sub SetError(group As String, errorType As String, quantity As Integer)
    
End Sub


Sub AddError(errorName As String, thrower As String, message As String, messageLoc As Range, Optional isWarn As Boolean = False)
    Dim ErrorLogCell As Range
    Set ErrorLogCell = ThisWorkbook.Worksheets("Errors").Range(errorName)
    
    If isWarn Then
        Set ErrorDashCell = ErrorLogCell.Offset(0, 2)
    Else
        Set ErrorDashCell = ErrorLogCell.Offset(0, 1)
    End If
    ErrorDashCell.Value = ErrorDashCell.Value + 1

    
    messageLoc.Value = AddStrings(messageLoc.Value, message)
    
    nextrow = Find_NextTopLevelRow(ErrorLogCell)
    
    ThisWorkbook.Worksheets("Errors").Rows(nextrow).Insert Shift:=xlDown
    ' Above row isn't included in With statement; With statement refers to newly inserted row
    With ThisWorkbook.Worksheets("Errors").Rows(nextrow)
        .ClearFormats
        If .OutlineLevel = 1 Then .group
        .IndentLevel = 2
    End With

    ThisWorkbook.Worksheets("Errors").Range("B" & nextrow).Value = thrower
End Sub


Sub Button_ShowErrorDetails()
    Dim RowIsHidden As Boolean

    For i = 6 To ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).row
        If Rows(i).EntireRow.Hidden = True Then
            RowIsHidden = True
            Exit For
        End If
    Next i

    If RowIsHidden Then
        ActiveSheet.Outline.ShowLevels RowLevels:=2
    Else
        ActiveSheet.Outline.ShowLevels RowLevels:=1
    End If
End Sub
