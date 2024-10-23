Attribute VB_Name = "Testing"
Sub Create_New_Release()
    'SelectFolder
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    Dim FldrPicker As FileDialog
    Dim myFolder As String
    
    'Have User Select Folder to Save to with Dialog Box
    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
    With FldrPicker
        .title = "Select Folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub 'Check if user clicked cancel button
        myFolder = .SelectedItems(1) & "\"
    End With
    'End SelectFolder
    
    'Ask the user to input a version number, suggest the last one in the changelog
    lastChange = ThisWorkbook.Sheets("Changelog").Range("A1").End(xlDown).Value
    inputVer = InputBox("Please provide version number", "Input Version Number", lastChange)
    
    'Clear this workbook
    Call Clear_All
    
    'Copy this workbook to a new file
    ThisWorkbook.SaveCopyAs myFolder & "QC Dashboard v" & inputVer & ".xlsm"
    Set wb = Workbooks.Open(myFolder & "QC Dashboard v" & inputVer & ".xlsm")
    
    'Hide the changelog, delete the testing worksheets and macros
    wb.Worksheets("Changelog").Visible = xlSheetHidden
    Application.DisplayAlerts = False
    wb.Worksheets("Testing").Delete
    wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents("Testing")
    Application.DisplayAlerts = True
    ThisWorkbook.Worksheets("Testing").Activate
    wb.Worksheets(1).Activate
    Call DeleteBrokenNamedRanges
    wb.Save
End Sub


Private Sub Button_TestProgressBar()

    ProgressBar.Show vbModeless
 
    Dim time1, time2
    Dim step As Integer
    Dim totalSteps As Integer
    step = 0
    totalSteps = 7
    Do Until step > totalSteps
        Call ProgressBarStep(step, totalSteps, "Testing")
        DoEvents
        Application.Wait (Now + TimeValue("00:00:01"))
        step = step + 1
    Loop
    Call ProgressBarSet(1, "Complete!")

End Sub


'Sums value from CountCol depending on if corresponding value in NameCol matches *any* of the filter strings provided
Function FilterSum(ByVal NameCol As Range, ByVal CountCol As Range, ParamArray filters() As Variant) As Variant
    'Build formula-syntax string array for filter parameters
    FilterArray = "{"
    For Each f In filters
        FilterArray = FilterArray & """" & f & ""","
    Next
    
    'Remove trailing comma and end the array with a right brace
    FilterArray = Left(FilterArray, Len(FilterArray) - 1)
    FilterArray = FilterArray & "}"

    'Extract the full addresses with sheet names from the ranges
    NameAddress = "'" & NameCol.Worksheet.Name & "'!" & NameCol.Address
    CountAddress = "'" & CountCol.Worksheet.Name & "'!" & CountCol.Address
    
    'Execute and return the formula, which does the following:
    ' - Uses SEARCH to generate a matrix comparing each value in NameCol to each of the filter strings
    ' - Uses MMULT to do an effective OR operation on each row, resulting in a TRUE value if any of the filter strings matched the value from NameCol
    ' - Passes the result to FILTER, which returns any values from CountCol corresponding to the passing NameCol values
    ' - Uses SUM to add the values that are passed through by FILTER
    FilterSum = Evaluate("=SUM(FILTER(" & CountAddress & ", MMULT(IF(ISNUMBER(SEARCH(" & FilterArray & ", " & NameAddress & ")), 1, 0), SEQUENCE(COLUMNS(" & FilterArray & "), 1, 1, 0)), 0))")
End Function

Sub Button_TestPathPrune()
    pathKMZ = ThisWorkbook.Sheets("File Imports").[Path_KMZ_Report].Value
    If pathKMZ = "" Then
        ThisWorkbook.Sheets("File Imports").Activate
        [Path_KMZ_Report].Select
        MsgBox ("path_KMZ_Report is not set. Please select a file, then try again.")
        End
    End If
    tempPath = Left(pathKMZ, InStrRev(pathKMZ, "\"))
    MsgBox tempPath
End Sub

Sub Button_TestSaveAsDialog()
    filePath = Application.GetSaveAsFilename( _
        fileFilter:="Text Files (*.txt), *.txt")
    If filePath <> False Then
        fileFullTitle = Mid(filePath, InStrRev(filePath, "\") + 1)
        fileTitle = Left(fileFullTitle, Len(fileFullTitle) - 4)
        MsgBox "Full path: " & filePath & vbNewLine & "File name with extension: " & fileFullTitle & vbNewLine & "File name: " & fileTitle
    End If
End Sub

Sub Clear_All()
    Call Clear_BOMDashboard
    Call Clear_KMZ
    Call Clear_SpliceReportDashboard
    Call Clear_TraceAddresses
    Call Clear_Imports
    Call Clear_ErrorDashboard
End Sub

Sub CheckOutlineLevel()
    MsgBox ActiveCell.Rows.OutlineLevel
End Sub

Sub CheckLastRow()
    MsgBox ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).row
End Sub


Sub DeleteBrokenNamedRanges()
Dim NR As Name
Dim numberDeleted As Variant

numberDeleted = 0
For Each NR In ActiveWorkbook.Names
    If InStr(NR.Value, "#REF!") > 0 Then
        NR.Delete
        numberDeleted = numberDeleted + 1
    End If
Next

MsgBox ("A total of " & numberDeleted & " broken Named Ranges deleted!")

End Sub

Sub Create_QC_Folder()
    'SelectFolder
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    Dim FldrPicker As FileDialog
    Dim myFolder As String
    
    'Have User Select Folder to Save to with Dialog Box
    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
    With FldrPicker
        .title = "Select Your Project Workspace"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub 'Check if user clicked cancel button
        myFolder = .SelectedItems(1) & "\"
    End With
    'End SelectFolder
    
    'Ask the user to input a version number, suggest the last one in the changelog
    lastChange = ThisWorkbook.Sheets("Changelog").Range("A1").End(xlDown).Value
    inputVer = InputBox("Please provide version number", "Input Version Number", lastChange)
    
    'Clear this workbook
    Call Clear_All
    
    Dim InputOLT As String
    InputOLT = InputBox("Please provide the name of your OLT", "Input OLT Name", "OLTNAME")
    
    CreateFolderPath (myFolder & InputOLT & "\Deliverables")
    CreateFolderPath (myFolder & InputOLT & "\Prism Docs")
    CreateFolderPath (myFolder & InputOLT & "\Reports")
    
    'Copy this workbook to a new file
    ThisWorkbook.SaveCopyAs myFolder & InputOLT & "\QC Dashboard v" & inputVer & " - " & InputOLT & ".xlsm"
    Set wb = Workbooks.Open(myFolder & InputOLT & "\QC Dashboard v" & inputVer & " - " & InputOLT & ".xlsm")
    
    'Hide the changelog, delete the testing worksheets and macros
    wb.Worksheets("Changelog").Visible = xlSheetHidden
    Application.DisplayAlerts = False
    wb.Worksheets("Testing").Delete
    wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents("Testing")
    Application.DisplayAlerts = True
    ThisWorkbook.Worksheets("Testing").Activate
    wb.Worksheets(1).Activate
    Call DeleteBrokenNamedRanges
    wb.Save
End Sub

