Attribute VB_Name = "SpliceReport"
Sub SortSpliceReport()
    Set imports = ThisWorkbook.Sheets("File Imports")
    Set Report = OpenPath(imports.[Path_Splice_Report])
    Report.Activate
    Call SortTabs
End Sub

Sub Clear_SpliceReportDashboard()
    ThisWorkbook.Worksheets("Splice Report").Range("6:" & Rows.Count).Clear
End Sub

Sub CheckAllErrors_SpliceReport()
    Dim imports As Worksheet
    Dim Dash As Worksheet
    Dim DashCR As Integer
    Dim Report As Workbook
    Dim ws As Worksheet
    Dim EqptName As String
    Dim DevName As String
    Dim SheathName As String
    Dim CR As Integer 'Current Row
    Dim SheathUUID As String
    Dim DevUUID As String
    Dim SheathDict As Variant
    Dim HasNAs As Boolean
    Dim ErrorCell As Range
    Dim WarnCell As Range
    Dim EqptCell As Range
    Dim SpliceConcat As String
    Dim IgnoreNaming As Boolean

    WasReportOpen = False
    Set imports = ThisWorkbook.Sheets("File Imports")
    Set Dash = ThisWorkbook.Sheets("Splice Report")
    If IsWorkBookOpen(imports.[Path_Splice_Report]) Then WasReportOpen = True
    Set Report = OpenPath(imports.[Path_Splice_Report])
    If Dash.CheckBoxes("Checkbox_IgnoreNaming").Value = 1 Then IgnoreNaming = True Else IgnoreNaming = False
    Dash.Activate
    
    ' Clear the dashboard
    Call Clear_SpliceReportDashboard
    Call Clear_ErrorDashboard("Splice Report")
    
    MaxRows = 10000 ' Define the max number of rows the macro can handle
    DashCR = 6
    Set SheathDict = CreateObject("Scripting.Dictionary")
    Set DeviceDict = CreateObject("Scripting.Dictionary")
    Set SpliceDict = CreateObject("Scripting.Dictionary")
    
    ' Loop through all Splice Report tabs
    For i = 1 To Report.Worksheets.Count
        Set ws = Report.Worksheets(i)
        
        EqptName = ws.Range("B1").Value
        Set EqptCell = Dash.Range("A" & DashCR)
        EqptCell.Value = EqptName
        Call AddWorksheetHyperlink(EqptCell, ws)
        If EqptName Like "*-*-*-*-*" Then
            If Not IgnoreNaming Then Call AddError("Error_Naming_Unnamed", EqptName, "Equipment is not named", Dash.Cells(DashCR, "D"))
        ElseIf Not ws.Name = EqptName Then
            If Not IgnoreNaming Then Call AddError("Error_Naming_SpliceReportTab", EqptName, "Equipment doesn't match tab name (possible duplicate name)", Dash.Cells(DashCR, "D"))
        End If
        
        ' Find the last row in the Sheaths section. May need a future revisit
        LR = ws.Range("A" & MaxRows).End(xlUp).row
        
        ' Skip this tab if there are no sheaths listed
        If (LR > MaxRows) Or (Not ws.Range("A6").Value = "SHEATH UUID") Then
            Dash.Cells(DashCR, "D").Value = "Unexpected format; Equipment is likely disconnected from sheaths"
            Dash.Cells(DashCR, "D").Font.Color = vbRed
            DashCR = DashCR + 1
        Else
            SheathDict.RemoveAll
            DeviceDict.RemoveAll
            SpliceDict.RemoveAll
            IsMST = False
    
            ' Check for Devices
            Set fnd = ws.Range("A1:A" & MaxRows).Find("OPTICAL SPLITTERS", , LookIn:=xlValues, LookAt:=xlWhole)
            If fnd Is Nothing Then
                LR = ws.Range("A" & MaxRows).End(xlUp).row
            Else
                LR = fnd.End(xlUp).row
                ' Device Error Checking
                DevFR = fnd.row + 2
                DevLR = ws.Range("A" & DevFR).End(xlDown).row
                ' Loop through the Device UUID column
                For Each cell In ws.Range("A" & DevFR & ":A" & DevLR)
                    DevUUID = cell.Value
                    ' Check if this row's device UUID is already in the collection
                    If (Not DevUUID = "") And (DeviceDict.Exists(DevUUID) = False) Then
                        ' Track the device UUID and the first row it was found on
                        DeviceDict.add DevUUID, cell.row
                    End If
                Next cell
                
                ' Error checking for each device
                For Each Key In DeviceDict.Keys
                    ' Reset per-device checks
                    HasConn = False
                    BadDevSplice = False
                    
                    Set ErrorCell = Dash.Range("D" & DashCR)
                    Set WarnCell = Dash.Range("E" & DashCR)
                    
                    FirstRow = CInt(DeviceDict(Key))
                    DevName = ws.Cells(FirstRow, "B").Value
                    Dash.Cells(DashCR, "B") = DevName
                    Dash.Cells(DashCR, "C") = "(Internal)"
                    ' Find the last row for this device
                    For row = FirstRow To DevLR
                        ' If we made it to the last device row in the tab, the last row for this device is the last row overall
                        If row = DevLR Then
                            LastRow = DevLR
                        ' Otherwise, detect whether a new device UUID was found, and if so, the previous row is the last row for this device
                        ElseIf (Not ws.Range("A" & row).Value = "") And (Not ws.Range("A" & row).Value = Key) Then
                            LastRow = row - 1
                            Exit For
                        End If
                    Next row
                    
                    For j = FirstRow To LastRow
                        SpliceType = ws.Range("E" & j).Value
                        If (Not SpliceType = "X") And (Not SpliceType = "") Then
                            ' Log that this device has a connection if the connection column isn't "X"
                            If Not HasConn Then HasConn = True
                            ' Check for double-splices
                            ' Concat Port name, device UUID, fiber, buffer, sheath UUID
                            SpliceConcat = ws.Range("G" & j) & ws.Range("I" & j) & ws.Range("J" & j) & ws.Range("K" & j) & ws.Range("O" & j)
                            If SpliceDict.Exists(SpliceConcat) = True Then
                                DoubSplice1 = SpliceDict(SpliceConcat)
                                DoubSplice2 = EqptName & ": " & DevName & "/" & ws.Range("C" & j)
                                Call AddError("Error_DoubleSplice", EqptName & "; " & DoubSplice1 & " & " & DoubSplice2, "DOUBLE SPLICED (see Errors tab for details)", ErrorCell)
                            Else
                                SpliceDict.add SpliceConcat, EqptName & ": " & DevName & "/" & ws.Range("C" & j)
                            End If
                        End If
                        
                        If BadDevSplice = False Then
                            If SpliceType = "<- CONTINUOUS ->" Or SpliceType = "<- N/A ->" Then
                                BadDevSplice = True
                                Call AddError("Error_Device_NonFusion", DevName, "Internal device has non-Fusion splice type", ErrorCell)
                            End If
                        End If
                    Next j
                    
                    If Not HasConn Then
                        Call AddError("Error_Disconn_Device", DevName, "Internal device is disconnected", ErrorCell)
                    End If
                    
                    
                    If Not ErrorCell.Value = "" Then
                        ErrorCell.Font.Color = vbRed
                    End If
                    If Not WarnCell.Value = "" Then
                        WarnCell.Font.ColorIndex = 44
                    End If
                    DashCR = DashCR + 1
                Next Key
            End If
    
            ' Loop through the Sheath UUID column
            For Each cell In ws.Range("A7:A" & LR)
                SheathUUID = cell.Value
                ' Check if this row's sheath UUID is already in the collection
                If (Not SheathUUID = "") And (SheathDict.Exists(SheathUUID) = False) Then
                    ' Track the sheath UUID and the first row it was found on
                    SheathDict.add SheathUUID, cell.row
                End If
            Next cell
            
            ' Attempt to detect whether this piece of equipment is an MST
            ' The splice report doesn't contain model information directly
            ' This check will give incorrect results if the attached sheath name is formatted incorrectly
            If (SheathDict.Count = 1) And (ws.Cells(7, "B").Value Like "*CT_*") Then IsMST = True
            
            
            'WIP IMPROVE CONTINUOUS FIBER COUNT CHECK
            'The idea here is to create an array which contains the sheath UUID, first row, last row, and fiber count for each sheath
            
            'Dim officeCounts(SheathDict.Count, 4) As String
            'For Each Key In SheathDict.Keys
                
            'Next Key
            
            
            ' Error checking for each sheath
            For Each Key In SheathDict.Keys
                ' Reset per-sheath checks
                HasNAs = False
                ContColorErr = False
                ContCountErr = False
                HasConn = False
                MSTContErr = False
                
                Set ErrorCell = Dash.Range("D" & DashCR)
                Set WarnCell = Dash.Range("E" & DashCR)
                
                FirstRow = CInt(SheathDict(Key))
                ' Find the last row for this sheath
                For row = FirstRow To LR
                    ' If we made it to the last sheath row in the tab, the last row for this sheath is the last row overall
                    If row = LR Then
                        LastRow = LR
                    ' Otherwise, detect whether a new sheath UUID was found, and if so, the previous row is the last row for this sheath
                    ElseIf (Not ws.Range("A" & row).Value = "") And (Not ws.Range("A" & row).Value = Key) Then
                        LastRow = row - 1
                        Exit For
                    End If
                Next row
                
                ' Log the sheath name to the Dashboard
                SheathName = ws.Cells(FirstRow, "B").Value
                Dash.Cells(DashCR, "B") = SheathName
                ' Log the Next Equipment name to the Dashboard
                If CleanString(ws.Cells(FirstRow, "C").Value) = CleanString(ws.Range("B1").Value) Then
                    NextEqptName = ws.Cells(FirstRow, "D").Value
                ElseIf CleanString(ws.Cells(FirstRow, "D").Value) = CleanString(ws.Range("B1").Value) Then
                    NextEqptName = ws.Cells(FirstRow, "C").Value
                Else
                    NextEqptName = "ERROR"
                End If
                Dash.Cells(DashCR, "C") = NextEqptName
                
                ' Find the fiber count for this sheath
                FiberCount = Application.WorksheetFunction.CountA(ws.Range("F" & FirstRow & ":F" & LastRow))
                
                ' Do one-time error checks
                If (Not SheathName Like "*" & FiberCount & "CT*") And (Not SheathName Like "*" & SheathName & "CT_") Then
                    If Not IgnoreNaming Then Call AddError("Error_Naming_FiberCT_Mismatch", SheathName, "Fiber name doesn't match Fiber Count", ErrorCell)
                End If
                ' Check for unnecessary spacing in sheath name
                If (SheathName Like " *") Or (SheathName Like "* ") Or (SheathName Like "*  *") Or (InStr(SheathName, vbLf) > 0) Then
                    If Not IgnoreNaming Then Call AddError("Error_Naming_SheathNameSpacing", EqptName, "Sheath name has unnecessary spacing", WarnCell, True)
                    ' Remove all extra spaces, line breaks, and other non-printable characters for future checks
                    SheathName = CleanString(SheathName)
                End If
                ' Check for unnecessary spacing in equipment name
                If (EqptName Like " *") Or (EqptName Like "* ") Or (EqptName Like "*  *") Or (InStr(EqptName, vbLf) > 0) Then
                    If Not IgnoreNaming Then Call AddError("Error_Naming_EqptNameSpacing", EqptName, "Equipment name has unnecessary spacing", WarnCell, True)
                    ' Remove all extra spaces, line breaks, and other non-printable characters for future checks
                    EqptName = CleanString(EqptName)
                End If
                ' Check for unnecessary spacing in connected equipment name
                If (NextEqptName Like " *") Or (NextEqptName Like "* ") Or (NextEqptName Like "*  *") Or (InStr(NextEqptName, vbLf) > 0) Then
                    If Not IgnoreNaming Then WarnCell.Value = AddStrings(WarnCell.Value, "Connected equipment name has unnecessary spacing")
                    ' Remove all extra spaces, line breaks, and other non-printable characters for future checks
                    NextEqptName = CleanString(NextEqptName)
                End If
                ' Check for sheath name formatting
                If Not (SheathName Like "*CT * TO *" Or SheathName Like "*CT_*") Then
                    If Not IgnoreNaming Then Call AddError("Error_Naming_Fiber_BadFormat", SheathName, "Sheath name is formatted incorrectly", ErrorCell)
                End If
                ' Check that sheath name contains both equipment names (or one equipment name for tails)
                If SheathName Like "*CT " & EqptName & " TO " & NextEqptName Then
                ElseIf SheathName Like "*CT " & NextEqptName & " TO " & EqptName Then
                ElseIf SheathName Like "*CT_" & EqptName Then
                ElseIf SheathName Like "*CT_" & NextEqptName Then
                Else
                    If Not IgnoreNaming Then Call AddError("Error_Naming_Fiber_NameMismatch", SheathName, "Sheath name doesn't match attached equipment", ErrorCell)
                End If
                
                ' DEBUG
                ' Dash.Cells(DashCR, "F") = FiberCount
                ' Dash.Cells(DashCR, "G") = FirstRow
                ' Dash.Cells(DashCR, "H") = LastRow
                ' Dash.Cells(DashCR, "I") = LR
                
                ' Loop through the rows in this sheath to check for errors
                For CR = FirstRow To LastRow
                    SpliceType = ws.Cells(CR, "J").Value
                    ' Skip rows with no splice type (usually extra rows for circuit info)
                    If Not SpliceType = "" Then
                        ' Check for NA splices
                        If (SpliceType = "<- N/A ->") And (HasNAs = False) Then
                            HasNAs = True
                            Call AddError("Error_NA_Splice_Type", EqptName & " (row " & CR & ")", "Connection has N/A splice on row " & CR, ErrorCell)
                        End If
                        ' Check for improper Continuous splices (not including MST splices)
                        If (Not IsMST) And (SpliceType = "<- CONTINUOUS ->") Then
                            ' Check for color-changing Continuous splices
                            If (ContColorErr = False) Then
                                If Not (ws.Cells(CR, "E").Value = ws.Cells(CR, "L").Value And ws.Cells(CR, "F").Value = ws.Cells(CR, "K").Value) Then
                                    ContColorErr = True
                                    Call AddError("Error_Cont_Color_Change", EqptName & " (row " & CR & ")", "Connection has color-changing Continuous splice on row " & CR, ErrorCell)
                                End If
                            End If
                            ' Check for Continuous splices between fibers of different counts (relies on the other fiber being properly named)
                            If (Not IgnoreNaming) And (ContCountError = False) Then
                                If (Not ws.Cells(CR, "O").Value Like "*" & FiberCount & "CT*") Then
                                    ContCountErr = True
                                    Call AddError("Error_Cont_Mismatch", EqptName & " (row " & CR & ")", "Uneven fiber counts are Continuous spliced on row " & CR, ErrorCell)
                                End If
                            End If
                        End If

                        If Not SpliceType = "X" Then
                            ' Check for a connection - if this is never triggered by the end of the sheath, then it's an un-spliced sheath
                            If Not HasConn Then HasConn = True
                            ' Check for double-splices
                            
                            ' DEPRECATED BY MAGELLAN UPDATE :)
                            ' Some splice reports have the right-hand columns shifted by one (not sure why this happens yet)
                            ' This IF statement figures out whether this splice report has shifted columns
                            ' If (Not IsEmpty(ws.Range("P" & CR))) And (Not IsEmpty(ws.Range("R" & CR))) Then
                            '     SpliceConcat = ws.Range("P" & CR) & ws.Range("R" & CR) & ws.Range("S" & CR) & ws.Range("K" & CR) & ws.Range("L" & CR) & ws.Range("O" & CR)
                            ' Else
                            '     SpliceConcat = ws.Range("Q" & CR) & ws.Range("R" & CR) & ws.Range("S" & CR) & ws.Range("K" & CR) & ws.Range("L" & CR) & ws.Range("O" & CR)
                            ' End If
                            
                            ' Concat Port name, device UUID, fiber, buffer, sheath UUID
                            SpliceConcat = ws.Range("Q" & CR) & ws.Range("T" & CR) & ws.Range("K" & CR) & ws.Range("L" & CR) & ws.Range("P" & CR)
                            
                            If SpliceDict.Exists(SpliceConcat) = True Then
                                DoubSplice1 = SpliceDict(SpliceConcat)
                                DoubSplice2 = SheathName & ": " & ws.Cells(CR, "E") & "/" & ws.Cells(CR, "F")
                                Call AddError("Error_DoubleSplice", EqptName & "; " & DoubSplice1 & " & " & DoubSplice2, "DOUBLE SPLICED (see Errors tab for details)", ErrorCell)
                            Else
                                SpliceDict.add SpliceConcat, SheathName & ": " & ws.Cells(CR, "E") & "/" & ws.Cells(CR, "F")
                            End If
                        End If
                        ' MST Error checks
                        If IsMST Then
                            If Not SpliceType = "<- CONTINUOUS ->" Then
                                MSTContErr = True
                            End If
                        End If
                    End If
                Next CR
                
                ' Log per-sheath errors if any were found
                If Not HasConn Then Call AddError("Error_Disconn_Sheath", EqptName, "Sheath is disconnected", ErrorCell)
                If MSTContErr Then Call AddError("Error_MST_NonCont", EqptName, "MST is not fully spliced as Continuous", ErrorCell)
                
                If Not ErrorCell.Value = "" Then
                    ErrorCell.Font.Color = vbRed
                End If
                If Not WarnCell.Value = "" Then
                    WarnCell.Font.ColorIndex = 44
                End If
                
                ' Move to the next row in the dashboard
                DashCR = DashCR + 1
            Next Key
        End If
        
        DashCR = DashCR + 1
    Next i
    
    If WasReportOpen = False Then Report.Close
    Dash.Columns("A:E").AutoFit
End Sub
