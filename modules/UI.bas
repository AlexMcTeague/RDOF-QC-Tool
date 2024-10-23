Attribute VB_Name = "UI"
Sub ProgressBarStep(thisStep As Integer, totalSteps As Integer, caption As String)

    ProgressBar.Text.caption = "Step " & thisStep & " of " & totalSteps & ": " & caption
    ProgressBar.Bar.Width = Round((thisStep / (totalSteps + 1)) * 200, 0)
    DoEvents

End Sub

Sub ProgressBarSet(completion As Single, caption As String)
    'completion should be a decimal value between 0 and 1
    
    ProgressBar.Text.caption = caption
    ProgressBar.Bar.Width = Round(completion * 200, 0)
    DoEvents

End Sub

Sub Button_SelectAllDeliverables()
'Imports all deliverable filepaths in a folder to the File Imports tab
    Dim folderPath As String
    Dim fileNames As Variant
    Dim ws As Worksheet
    Dim foundIndex As Integer
    Dim Dict As Variant
    
    'End if user didn't select a folder
    folderPath = SelectFolder()
    If folderPath = "" Then End
    
    'End if selected folder is empty
    fileNames = listfiles(folderPath)
    If IsEmpty(fileNames) Then End
    
    Set ws = ThisWorkbook.Sheets("File Imports")
    Set Dict = CreateObject("Scripting.Dictionary")
    
    Dict.add "Path_Before_Print", "*BEFORE_FIBER*"
    Dict.add "Path_After_Print", "*AFTER_FIBER*"
    Dict.add "Path_Overview_Print", "*OVERVIEW_FIBER*"
    Dict.add "Path_Grid_Print", "*GRID_FIBER*"
    Dict.add "Path_BOMs", "*_BOM*" 'Some people use BOM_Report, some use just BOM
    Dict.add "Path_Overall_BOM", "*OVERALL*"
    Dict.add "Path_KMZ_Report", "*KMZ_FIBER*"
    Dict.add "Path_HAF", "*HAF*.xlsx"
    Dict.add "Path_HAF_CSV", "*HAF*.csv"
    Dict.add "Path_PON_Calc", "*PON*"
    Dict.add "Path_MOP", "*MOP*"
    Dict.add "Path_Splice_Report", "*SPLICE*"
    
    For Each Key In Dict.Keys()
        ws.Range(Key).ClearContents
        foundIndex = FindStringInArray(CStr(Dict(Key)), fileNames)
        If foundIndex > -1 Then
            ws.Range(Key) = folderPath + fileNames(foundIndex)
        End If
    Next Key
End Sub

Sub Clear_Imports()
    Set ws = ThisWorkbook.Sheets("File Imports")
    
    ws.[Path_Before_Print].ClearContents
    ws.[Path_After_Print].ClearContents
    ws.[Path_Overview_Print].ClearContents
    ws.[Path_Grid_Print].ClearContents
    ws.[Path_BOMs].ClearContents
    ws.[Path_Overall_BOM].ClearContents
    ws.[Path_KMZ_Report].ClearContents
    ws.[Path_HAF].ClearContents
    ws.[Path_HAF_CSV].ClearContents
    ws.[Path_PON_Calc].ClearContents
    ws.[Path_MOP].ClearContents
    ws.[Path_Splice_Report].ClearContents
    
    ws.[Path_SG_HAF].ClearContents
    ws.[Path_SG_HAF_CSV].ClearContents
    ws.[Path_OLT_FWD_Trace].ClearContents
    ws.[Path_OLT_FWD_Trace].Offset(1, 0).ClearContents
    ws.[Path_OLT_FWD_Trace].Offset(2, 0).ClearContents
    ws.[Path_OLT_FWD_Trace].Offset(3, 0).ClearContents
End Sub


Sub Trim_File_Dates()
    Dim path As String
    Dim edited As Integer
    
    edited = 0
    
    Set objRegEx = CreateObject("vbscript.regexp")
    objRegEx.Global = True
    objRegEx.IgnoreCase = True
    objRegEx.MultiLine = True

    
    MsgBox "This macro removes the dates that MQMS adds to the end of filenames." & vbNewLine & _
        "This macro will error if you have any of these files open, or if the trimmed name already exists."

    For Each cell In ActiveWorkbook.Worksheets("File Imports").Range("C4:C30")
        path = cell.Value
        
        If path Like "*_??-??-????_??-??-??.*" Then
            objRegEx.Pattern = "_..-..-...._..-..-.."
            If IsWorkBookOpen(path) Then
                wbName = Mid(path, InStrRev(path, "\") + 1)
                Workbooks(wbName).Close
            End If
            newPath = objRegEx.Replace(path, "")
            Name path As newPath
            cell.Value = newPath
            
            edited = edited + 1
        End If
    Next cell
    
    If edited > 0 Then MsgBox edited & " file names were edited!"
End Sub

