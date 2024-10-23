Attribute VB_Name = "GeneralUseMacros"
Sub ResumeUpdating()
'Use this to set ScreenUpdating back to True if a macro breaks after setting it to False
    
    Application.ScreenUpdating = True
    
End Sub

Sub UnHideAll()
'Unhides all sheets, including "very hidden" sheets

    For Each ws In ActiveWorkbook
        ws.Visible = xlSheetVisible
    Next
    
End Sub

Sub SortTabs()
    
    Application.ScreenUpdating = False
    
    Dim ShCount As Integer, i As Integer, j As Integer
    ShCount = ActiveWorkbook.Sheets.Count

    For i = 1 To ShCount - 1
        For j = i + 1 To ShCount
            If UCase(ActiveWorkbook.Sheets(j).Name) < UCase(ActiveWorkbook.Sheets(i).Name) Then
                ActiveWorkbook.Sheets(j).Move Before:=ActiveWorkbook.Sheets(i)
            End If
        Next j
    Next i

    Application.ScreenUpdating = True

End Sub

Sub FindReplaceAllInRange(fnd As Variant, rplc As Variant, rng As Range)

    rng.Cells.Replace What:=fnd, Replacement:=rplc, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False

End Sub

Function Button_AdjacentCell(Optional dir As String = "R") As Range

    Dim cellTarget As Range
    Dim result As Range
    
    Set cellTarget = ActiveSheet.Buttons(Application.Caller).TopLeftCell
    
    Select Case LCase(Left(dir, 1))
        Case "l"
            Set result = cellTarget.Offset(0, -1)
        Case "r"
            Set result = cellTarget.Offset(0, 1)
        Case "u"
            Set result = cellTarget.Offset(-1, 0)
        Case "d"
            Set result = cellTarget.Offset(1, 0)
        Case Else
            MsgBox ("Unrecognized button argument. Ask the maintainer of this spreadsheet for assistance.")
            Set result = cellTarget.Offset(0, 1)
    End Select
    
    Set Button_AdjacentCell = result

End Function

Sub Button_CopyToClipboard(Optional dir As String = "R")

    Dim objCP As Object
    Dim cellTarget As Range
    
    Set objCP = CreateObject("HtmlFile")
    Set cellTarget = Button_AdjacentCell(dir)
    
    objCP.ParentWindow.ClipboardData.SetData "text", CStr(cellTarget.Value)
    
End Sub

Sub Button_SelectFilepath(Optional dir As String = "R", Optional multiSelect As Boolean = False)

    Dim cellTarget As Range
    Set cellTarget = Button_AdjacentCell(dir)

    If multiSelect = True Then
        ' https://stackoverflow.com/questions/50382575/prompt-user-to-select-multiple-files-and-perform-the-same-action-on-all-files
        Dim fDialog As FileDialog
        Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
        With fDialog
            .AllowMultiSelect = True
            .title = "Select Files"
            .filters.Clear
            .filters.add "All Files (*.*)", "*.*"
    
            If .Show = True Then
                Dim fPath As Variant
                For Each fPath In .SelectedItems
                    cellTarget.Value = fPath
                    Set cellTarget = cellTarget.Offset(1, 0)
                Next
            End If
        End With
    Else
        cellTarget.Value = Application.GetOpenFilename(fileFilter:="All Files (*.*), *.*", title:="Select A File")
    End If

End Sub


Sub Button_SelectMultiFilepath(Optional dir As String = "R")

    Call Button_SelectFilepath(dir, True)
    
End Sub

Sub Button_OpenFilepath(Optional dir As String = "R")

    Dim filePath As String
    
    filePath = CStr(Button_AdjacentCell(dir).Value)
    If (Not filePath = "") And (Not filePath = "False") Then
        fileName = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))
        For Each wb In Application.Workbooks()
            If wb.Name = fileName Then
                If GetLocalPath(wb.FullName) <> filePath Then
                    ThisWorkbook.Sheets("File Imports").Activate
                    MsgBox ("Excel can't open two workbooks with the same name at the same time." & vbNewLine & "Select a different file, or close the other workbook named " & fileName)
                    End
                Else
                    wb.Activate
                    Exit Sub
                End If
            End If
        Next
    
        ThisWorkbook.FollowHyperlink filePath
    End If
End Sub

Sub Button_OpenNamedImport(Name As String)

    Dim wb As Workbook
    Dim cell As Range
    
    Set cell = ThisWorkbook.Sheets("File Imports").Range(Name)
    Set wb = OpenPath(cell)
End Sub

Public Function OpenPath(cell As Range) As Workbook
    Dim wb As Workbook

    ' If there isn't a path available, end all macros and inform the user
    path = cell.Value
    If IsEmpty(cell) Or path = "" Or path = False Then
        ThisWorkbook.Sheets("File Imports").Activate
        cell.Select
        MsgBox (cell.Name.Name & " is not set. Please select a file, then try again.")
        End
    Else
        ' If the file can't be found, end all macros and inform the user
        If dir(path) = "" Then
            ThisWorkbook.Sheets("File Imports").Activate
            cell.Select
            MsgBox ("File doesn't exist at " & path & ". Please select a different file, then try again.")
            End
        End If
        ' If a file with an identical name (but a different path) is already open, end all macros and inform the user
        fileName = Right(path, Len(path) - InStrRev(path, "\"))
        For Each wb In Application.Workbooks()
            If wb.Name = fileName Then
                If GetLocalPath(wb.FullName) <> path Then
                    ThisWorkbook.Sheets("File Imports").Activate
                    cell.Select
                    MsgBox ("Excel can't open two workbooks with the same name at the same time." & vbNewLine & "Select a different file, or close the other workbook named " & fileName)
                    End
                End If
            End If
        Next
        
        Set wb = Workbooks.Open(path, UpdateLinks:=0)
        Set OpenPath = wb
    End If
End Function

Public Function SelectFolder() As String
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    Dim FldrPicker As FileDialog
    Dim myFolder As String
    
    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
    'Have user select folder with Dialog Box
    With FldrPicker
        .title = "Select Folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function 'Check if user clicked cancel button
        myFolder = .SelectedItems(1) & "\"
    End With
    SelectFolder = myFolder 'Returns the path to the folder as a String
End Function

Function listfiles(ByVal path As String) As Variant
    'SOURCE: https://stackoverflow.com/questions/66441427/get-all-file-names-in-folder-to-array-and-sort-in-alphabetically-with-string-and

    Dim vaArray     As Variant
    Dim oFile       As Object
    Dim oFiles      As Object

    Set oFiles = CreateObject("Scripting.FileSystemObject").GetFolder(path).Files

    If oFiles.Count = 0 Then Exit Function

    ReDim vaArray(1 To oFiles.Count)
    i = 1
    For Each oFile In oFiles
        vaArray(i) = oFile.Name
        i = i + 1
    Next

    listfiles = vaArray

End Function

Function FindStringInArray(stringToBeFound As String, arr As Variant) As Long
    'SOURCE: https://stackoverflow.com/questions/10951687/how-to-search-for-string-in-an-array
    
    'Default return value if value not found in array
    FindStringInArray = -1

    For i = LBound(arr) To UBound(arr)
    If arr(i) Like stringToBeFound Then
      FindStringInArray = i
      Exit For
    End If
  Next i
End Function

Sub CreateFolderPath(folderPath As String)
    'Create all the folders in a folder path
    'SOURCE: https://exceloffthegrid.com/vba-code-create-delete-manage-folders/
    Dim individualFolders() As String
    Dim tempFolderPath As String
    Dim arrayElement As Variant

    'Split the folder path into individual folder names
    individualFolders = Split(folderPath, "\")

    'Loop though each individual folder name
    For Each arrayElement In individualFolders

        'Build string of folder path
        tempFolderPath = tempFolderPath & arrayElement & "\"
 
        'If folder does not exist, then create it
        If dir(tempFolderPath, vbDirectory) = "" Then
 
            MkDir tempFolderPath
 
        End If
 
    Next arrayElement
    'End CreateFolders
End Sub

Function WorksheetExists(ByVal shtName As String, Optional wb As Workbook) As Boolean
'Source: https://stackoverflow.com/questions/6688131/test-or-check-if-sheet-exists
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Function AddStrings(orig As String, add As String) As String
    If orig = "" Then
        AddStrings = add
    Else
        result = orig & ", " & add
        AddStrings = result
    End If
End Function

Function Find_NextVisibleRow(row As Integer) As Integer
    i = row + 1
    Do While Rows(i).EntireRow.Hidden = True
        i = i + 1
    Loop
    Find_NextVisibleRow = i - 1
End Function

Function Find_NextTopLevelRow(cell As Range) As Integer
    i = cell.row + 1
    Do While cell.Worksheet.Rows(i).OutlineLevel > 1
        i = i + 1
    Loop
    Find_NextTopLevelRow = i
End Function

Sub AddWorksheetHyperlink(targetCell As Range, ws As Worksheet)
    ' targetCell is the cell where you want the link to appear
    ' ws is the sheet you want to link to
    linkTo = ws.Parent.FullName & "#'" & ws.Name & "'!A1"
    targetCell.Parent.Hyperlinks.add targetCell, linkTo
End Sub

Function IsWorkBookOpen(fileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open fileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False ' File exists, but is not open
    Case 53:   IsWorkBookOpen = False ' File Not Found
    Case 70:   IsWorkBookOpen = True  ' File is already open
    Case 75:   IsWorkBookOpen = False ' FileName is empty
    Case Else: Error ErrNo
    End Select
End Function

Sub UnloadAllForms()
'Unloads all open user forms
    Dim i As Integer
    For i = VBA.UserForms.Count - 1 To 0 Step -1
        Unload VBA.UserForms(i)
    Next
End Sub

Function CleanString(ByVal inputString As String) As String
    Dim result As String
    result = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(inputString))
    CleanString = result
End Function

Function MergeDicts(ParamArray dictionaries() As Variant) As Dictionary
' Inspiration: https://stackoverflow.com/questions/21903530/combine-2-scripting-dictionaries-or-collections-of-key-item-pairs
    Dim CombinedDict, Key, Dict
    Set CombinedDict = CreateObject("Scripting.Dictionary")
    Set Dict = CreateObject("Scripting.Dictionary")

    For i = 0 To UBound(dictionaries())
        Set Dict = dictionaries(i)
        For Each Key In Dict.Keys()
            CombinedDict.Item(Key) = Dict(Key)
        Next Key
    Next i

    Set MergeDicts = CombinedDict
End Function
