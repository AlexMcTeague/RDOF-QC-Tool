Attribute VB_Name = "Continuity"
Sub CheckAllErrors_TraceAddresses()
    Dim Dash As Worksheet
    Dim imports As Worksheet
    'Dim HAF As Worksheet
    'Dim SGHAF As Variant
    'Dim Trace As Worksheet
    Dim FoundCell As Range
    Dim ErrorCell As Range
    Dim WarnCell As Range
    Dim FirstDashRow As Integer
    Dim usedPort As String
    Dim pathTrace As Range
    Dim errorPortInfo As String
    
    Dim DashCoordCol As String: DashCoordCol = "A"
    Dim DashAddrCol As String: DashAddrCol = "B"
    Dim DashHAFEncCol As String: DashHAFEncCol = "C"
    Dim DashHAFPortCol As String: DashHAFPortCol = "D"
    Dim DashTrcEncCol As String: DashTrcEncCol = "E"
    Dim DashTrcPortCol As String: DashTrcPortCol = "F"
    Dim DashTraceNumCol As String: DashTraceNumCol = "G"
    Dim DashResultCol As String: DashResultCol = "H"

    
    FirstDashRow = 6 ' The row in the Port Continuity dashboard where the data begins
    
    'Set up the sheets for the calculations, check that the user selected a HAF file
    WasHAFOpen = False
    WasSGHAFOpen = False
    Set Dash = ThisWorkbook.Worksheets("Port Continuity")
    Set imports = ThisWorkbook.Worksheets("File Imports")
    If IsWorkBookOpen(imports.[Path_HAF]) Then WasHAFOpen = True
    
    Dim HAF As HAFSheet: Set HAF = New HAFSheet
    Dim SG_HAF As HAFSheet
    Set HAF.sheet = OpenPath(imports.[Path_HAF]).Sheets(1)
    Dim Trace As TRCSheet
    
    If Not imports.[Path_SG_HAF].Value = "" Then
        If IsWorkBookOpen(imports.[Path_SG_HAF]) Then WasSGHAFOpen = True
        Set SG_HAF = New HAFSheet
        Set SG_HAF.sheet = OpenPath(imports.[Path_SG_HAF]).Sheets(1)
    End If
    
    'Check that the user selected at least one forward trace file
    Dim allTracePaths As Variant
    For i = 0 To 3
        allTracePaths = allTracePaths & imports.[Path_OLT_FWD_Trace].Offset(i, 0).Value
    Next i
    If allTracePaths = "" Then
        imports.Activate
        imports.Range("Path_OLT_FWD_Trace").Select
        MsgBox ("No Forward Traces selected. Please select at least one file and try again.")
        End
    End If
    
    'Clear the dashboard
    Call Clear_TraceAddresses
    Call Clear_ErrorDashboard("Port Continuity")
    Dash.Columns(DashHAFPortCol).Hidden = False ' This prevents the macro from (for some reason) occasionally losing track of the values in hidden columns
    Dash.Columns(DashTrcPortCol).Hidden = False
    
    'Slice Setup
    Dim HAFSliceDict As Dictionary
    Set HAFSliceDict = CreateObject("Scripting.Dictionary")
    Dim HAFColArray As Variant
    HAFColArray = Array(HAF_LAT, HAF_LONG, HAF_HOUSE_NUMBER, HAF_STREET_NAME, HAF_STREET_NAME, HAF_STREET_TYPE, HAF_COMMENT, HAF_POLE_PORT_NUMS)
    
    'Loop through HAF(s) and craft slices
    Dim LR As Integer
    LR = HAF.sheet.Range(HAF.get_letter(HAF_HOUSE_NUMBER) & HAF.sheet.Rows.Count).End(xlUp).row
    i = 2
    Do While i <= LR
        HAFSliceDict.add i, HAF.slice(HAFColArray, i)
        i = i + 1
    Loop
    
    If Not (SG_HAF Is Nothing) Then
        LR = SG_HAF.sheet.Range(SG_HAF.get_letter(HAF_HOUSE_NUMBER) & SG_HAF.sheet.Rows.Count).End(xlUp).row
        i = 2
        Do While i <= LR
            HAFSliceDict.add "SG " & i, SG_HAF.slice(HAFColArray, i)
            i = i + 1
        Loop
    End If

    'Loop through the dictionary of slices to copy data to the dashboard
    DashRow = FirstDashRow
    For Each Key In HAFSliceDict.Keys()
        Set slice = HAFSliceDict(Key)
        Dash.Range(DashCoordCol & DashRow) = CStr(Round(CDbl(slice.Value(HAF_LAT)), 6)) & ", " & CStr(Round(CDbl(slice.Value(HAF_LONG)), 6))
        Dash.Range(DashAddrCol & DashRow) = slice.Value(HAF_HOUSE_NUMBER) & " " & slice.Value(HAF_STREET_NAME) & " " & slice.Value(HAF_STREET_TYPE)
        Dash.Range(DashHAFEncCol & DashRow) = slice.Value(HAF_COMMENT)
        portNum = Trim(Replace(Right(slice.Value(HAF_POLE_PORT_NUMS), 2), "T", ""))
        Dash.Range(DashHAFPortCol & DashRow) = slice.Value(HAF_COMMENT) & " (PORT " & portNum & ")"
        DashRow = DashRow + 1
    Next Key
    
    'Close the HAF and SG HAF if they were closed before the macro was run
    If WasHAFOpen = False Then HAF.sheet.Parent.Close
    If (Not SG_HAF Is Nothing) And (WasSGHAFOpen = False) Then SGHAF.sheet.Parent.Close
    
    'Import all enclosures from the trace(s) that have continuity to the OLT
    DashRow = FirstDashRow
    Dim TraceSplitUUIDs As Dictionary
    Set TraceSplitUUIDs = CreateObject("Scripting.Dictionary")
    Dim TraceTempDict As Dictionary
    Set TraceTempDict = CreateObject("Scripting.Dictionary")
    Dim TraceSplice As TRCSlice
    Dim TraceColArray As Variant
    TraceColArray = Array(TRC_ENC_UUID, TRC_PATH_SPLIT, TRC_DEVICE_UUID_L, TRC_PORT_NAME_R, TRC_ENC_TYPE, TRC_DEVICE_NAME_R, TRC_ENC_NAME)
    
    ' Loop through the (up to) four forward traces
    For traceNum = 1 To 4
        Set pathTrace = imports.[Path_OLT_FWD_Trace].Offset(traceNum - 1, 0)
        prevEncRow = TRC_HEADER_ROW
        If Not IsEmpty(pathTrace) Then
            Set Trace = New TRCSheet
            Set Trace.sheet = OpenPath(pathTrace).Sheets(1)
            
            ' Skip this document if there's no header row (empty trace)
            If Not Range("A" & TRC_HEADER_ROW).Value = "" Then
                LR = Trace.sheet.Range(Trace.get_letter(TRC_ENC_UUID) & Trace.sheet.Rows.Count).End(xlUp).row
                TraceSplitUUIDs.RemoveAll ' Clear the tracked device UUIDs
                
                For traceRow = TRC_HEADER_ROW + 1 To LR
                    ' We only really need info if we encounter a split in the trace, or on the last row of the trace
                    If Trace(TRC_PATH_SPLIT, traceRow) = "True" Then
                        Set slice = Trace.slice(TraceColArray, traceRow)
                        ' If we haven't seen this device UUID, that means we're splitting the current branch again.
                        If Not TraceSplitUUIDs.Exists(slice.Value(TRC_DEVICE_UUID_L)) Then
                            ' Add this row to the list, using the UUID as a key
                            TraceSplitUUIDs.add slice.Value(TRC_DEVICE_UUID_L), slice
                        ' If we've already seen this device UUID, then we're returning to the next out-port of that device.
                        Else
                            If Trace.Value(TRC_PATH_SPLIT, prevEncRow) = "True" Then
                                ' If the previous enclosure's row also contains an out-port, that port must not be spliced, so we can ignore it.
                            Else
                                Set prevEncSlice = Trace.slice(TraceColArray, prevEncRow)
                                If (prevEncSlice.Value(TRC_PORT_NAME_R) Like "PORT*") And (prevEncSlice.Value(TRC_ENC_TYPE) = "HYBRID") Then
                                    ' If the previous enclosure's row shows a Hybrid's Port#, then we log it to the Dashboard
                                    Dash.Range(DashTrcEncCol & DashRow).Value = prevEncSlice.Value(TRC_DEVICE_NAME_R)
                                    portNum = Right(prevEncSlice.Value(TRC_PORT_NAME_R), Len(prevEncSlice.Value(TRC_PORT_NAME_R)) - 4)
                                    Dash.Range(DashTrcPortCol & DashRow).Value = prevEncSlice.Value(TRC_ENC_NAME) & " (PORT " & portNum & ")"
                                Else
                                    ' If the previous enclosure's row doesn't show a Port#, then the previous split didn't end at a hybrid's port properly
                                    ' Since we haven't logged the new slice yet, we can use the last entry in the SplitUUID dictionary to get info about the previous slice
                                    Set oldSlice = TraceSplitUUIDs.Item(TraceSplitUUIDs.Keys(UBound(TraceSplitUUIDs.Keys)))
                                    errorPortInfo = oldSlice.Value(TRC_ENC_NAME) & ": " & oldSlice.Value(TRC_DEVICE_NAME_R) & ": " & oldSlice.Value(TRC_PORT_NAME_R)
                                    ' Now we can log the warning
                                    Set ErrorCell = Dash.Range(DashResultCol & DashRow)
                                    Call AddError("Error_Trace_Unnecessary_Split_Port", errorPortInfo, "[" & errorPortInfo & "] Splitter port doesn't trace to tap", ErrorCell, True)
                                    ErrorCell.Font.ColorIndex = 44
                                End If
                                Dash.Range(DashTraceNumCol & DashRow).Value = "Trace #" & traceNum
                                DashRow = DashRow + 1
                            End If
                            ' Now it's time to save the current slice to the dictionary, so the next loop can read it.
                            ' We only need to keep the slices "above" this slice in the structure
                            TraceTempDict.RemoveAll
                            For Each Key In TraceSplitUUIDs.Keys
                                ' We loop through the existing UUID dictionary...
                                If Not Key = slice.Value(TRC_DEVICE_UUID_L) Then
                                    ' ...keeping everything until we find the matching UUID
                                    TraceTempDict.add Key, TraceSplitUUIDs(Key)
                                Else
                                    ' When we find the matching UUID, we add our new slice and stop copying from the original dict, ensuring it's the latest entry
                                    TraceTempDict.add slice.Value(TRC_DEVICE_UUID_L), slice
                                    Exit For
                                End If
                            Next Key
                            ' Then we can copy the result back to the original dict to use it in the next loop
                            TraceSplitUUIDs.RemoveAll
                            For Each Key In TraceTempDict
                                TraceSplitUUIDs.add Key, TraceTempDict(Key)
                            Next Key
                        End If
                    ElseIf traceRow = LR Then
                        ' Special handling for the last row of the trace (if it's a new unspliced port, it will be handled and correctly ignored by the code above)
                        Set slice = Trace.slice(TraceColArray, traceRow)
                        If (slice.Value(TRC_PORT_NAME_R) Like "PORT*") And (slice.Value(TRC_ENC_TYPE) = "HYBRID") Then
                            ' If the last row shows a Hybrid's Port#, then we log it to the Dashboard
                            Dash.Range(DashTrcEncCol & DashRow).Value = slice.Value(TRC_DEVICE_NAME_R)
                            portNum = Right(slice.Value(TRC_PORT_NAME_R), Len(slice.Value(TRC_PORT_NAME_R)) - 4)
                            Dash.Range(DashTrcPortCol & DashRow).Value = slice.Value(TRC_ENC_NAME) & " (PORT " & portNum & ")"
                        Else
                            ' If the last row doesn't show a Port#, then the previous split didn't end at a hybrid's port properly
                            ' Since we haven't logged the new slice yet, we can use the last entry in the SplitUUID dictionary to get info about the previous slice
                            Set oldSlice = TraceSplitUUIDs.Item(TraceSplitUUIDs.Keys(UBound(TraceSplitUUIDs.Keys)))
                            errorPortInfo = oldSlice.Value(TRC_ENC_NAME) & ": " & oldSlice.Value(TRC_DEVICE_NAME_R) & ": " & oldSlice.Value(TRC_PORT_NAME_R)
                            ' Now we can log the warning
                            Set ErrorCell = Dash.Range(DashResultCol & DashRow)
                            Call AddError("Error_Trace_Unnecessary_Split_Port", errorPortInfo, "[" & errorPortInfo & "] Splitter port doesn't trace to tap", ErrorCell, True)
                            ErrorCell.Font.ColorIndex = 44
                        End If
                        Dash.Range(DashTraceNumCol & DashRow).Value = "Trace #" & traceNum
                        DashRow = DashRow + 1
                    End If
                    
                    
                    ' We keep track of the last row that has enclosure data, that way we can check it if we reach a new split
                    If Not (Trace.Value(TRC_ENC_UUID, traceRow) = "") Then
                        prevEncRow = traceRow
                    End If
                Next traceRow
            End If
            Trace.sheet.Parent.Close
        End If
    Next traceNum
    
    'Sort the data
    'Sort the left half (HAF) and right half (Trace) independently. Unassociated errors remain separated because blank lines in the Results column are considered in the right-half sort
    Dash.Range(DashCoordCol & FirstDashRow & ":" & DashHAFPortCol & FirstDashRow + 4096).Sort Key1:=Dash.Range(DashHAFEncCol & FirstDashRow), Order1:=xlAscending, Key2:=Dash.Range(DashHAFPortCol & FirstDashRow), Order2:=xlAscending, Header:=xlNo
    Dash.Range(DashTrcEncCol & FirstDashRow & ":" & DashResultCol & FirstDashRow + 4096).Sort Key1:=Dash.Range(DashTrcEncCol & FirstDashRow), Order1:=xlAscending, Key2:=Dash.Range(DashTrcPortCol & FirstDashRow), Order2:=xlAscending, Header:=xlNo
    LRHAF = WorksheetFunction.Min(Dash.Range(DashCoordCol & Dash.Rows.Count).End(xlUp).row, 4096)
    LRTRC = WorksheetFunction.Min(Dash.Range(DashTrcEncCol & Dash.Rows.Count).End(xlUp).row, 4096)
    
    ' Alert the user if no addresses/ports were found, or too many were found
    If (LRHAF = FirstDashRow - 1) Then
        imports.Activate
        imports.Range("Path_HAF").Select
        MsgBox "Could not find addresses in HAF. Make sure the correct file is selected"
        End
    End If
    If (LRTRC = FirstDashRow - 1) Then
        imports.Activate
        imports.Range("Path_OLT_FWD_Trace").Select
        MsgBox "Could not find ports in trace(s). Make sure the FWD Trace files have data"
        End
    End If
    If (LRHAF = 4096) Then
        imports.Activate
        imports.Range("Path_HAF").Select
        MsgBox "Import Error: Found too many addresses in HAF"
        End
    End If
    If (LRTRC = 4096) Then
        imports.Activate
        imports.Range("Path_OLT_FWD_Trace").Select
        MsgBox "Import Error: Found too many ports in traces"
        End
    End If
    
    'Analyze the data
    'i: incrementer to loop through HAF enclosure rows (in dashboard)
    i = FirstDashRow
    Do While i <= LRHAF
        Set CoordCell = Dash.Range(DashCoordCol & i)
        Set AddressCell = Dash.Range(DashAddrCol & i)
        Set HAFEncCell = Dash.Range(DashHAFEncCol & i)
        Set HAFPortCell = Dash.Range(DashHAFPortCol & i)
        Set TraceEncCell = Dash.Range(DashTrcEncCol & i)
        Set TracePortCell = Dash.Range(DashTrcPortCol & i)
        Set TraceNumCell = Dash.Range(DashTraceNumCol & i)
        Set ErrorCell = Dash.Range(DashResultCol & i)
        
        'If there are more addresses than trace enclosures, the only errors left to check are "address doesn't trace" and "address isn't linked"
        If i > LRTRC Then
            Range(TraceNumCell, ErrorCell).Insert Shift:=xlDown
            Set TraceNumCell = Dash.Range(DashTraceNumCol & i)
            Set ErrorCell = Dash.Range(DashResultCol & i)
            If HAFEncCell.Value = "" Then
                Call AddError("Error_Address_Unlinked", AddressCell.Value, "Address isn't linked to a named enclosure", ErrorCell)
            Else
                Call AddError("Error_Address_NoTrace", HAFPortCell.Value, HAFPortCell.Value & " doesn't trace", ErrorCell)
            End If
            ErrorCell.Font.Color = vbRed
        'The rest of the errors are checked below
        Else
            If HAFEncCell.Value = "" Then
                Call AddError("Error_Address_Unlinked", AddressCell.Value, "Address isn't linked to a named enclosure", ErrorCell)
                ErrorCell.Font.Color = vbRed
                Range(TraceEncCell, TraceNumCell).Insert Shift:=xlDown
                ErrorCell.Offset(1, 0).Insert Shift:=xlDown
                LRTRC = LRTRC + 1
            ElseIf HAFPortCell.Value = TracePortCell.Value Then
                ErrorCell.Value = "Trace Successful"
                ErrorCell.Font.Color = vbGreen
            Else
                Set FoundCell = Nothing
                Set FoundCell = Dash.Range(TracePortCell, Dash.Cells(LRTRC, TracePortCell.column)).Find(What:=HAFPortCell.Value, LookIn:=xlValues, LookAt:=xlWhole)
                If FoundCell Is Nothing Then
                    Call AddError("Error_Address_NoTrace", HAFPortCell.Value, HAFPortCell.Value & " doesn't trace", ErrorCell)
                    ErrorCell.Font.Color = vbRed
                    Range(TraceEncCell, TraceNumCell).Insert Shift:=xlDown
                    ErrorCell.Offset(1, 0).Insert Shift:=xlDown
                    LRTRC = LRTRC + 1
                Else
                    Call AddError("Error_Trace_NoAddressMatch", TracePortCell.Value, TracePortCell.Value & " traces but doesn't have matching address in HAF", ErrorCell)
                    ErrorCell.Font.Color = vbRed
                    Dash.Range(CoordCell, HAFPortCell).Insert Shift:=xlDown
                    LRHAF = LRHAF + 1
                End If
            End If
        End If
        
        i = i + 1
    Loop
    Do While i <= LRTRC
        Set ErrorCell = Dash.Range("H" & i)
        Call AddError("Error_Trace_NoAddressMatch", Dash.Range("F" & i).Value, Dash.Range("F" & i).Value & " traces but doesn't have matching address in HAF", ErrorCell)
        ErrorCell.Font.Color = vbRed
        i = i + 1
    Loop
    
    'Fit columns to match data
    Dash.Columns(DashCoordCol & ":" & DashResultCol).AutoFit
    Dash.Columns(DashHAFPortCol).Hidden = True
    Dash.Columns(DashTrcPortCol).Hidden = True
    Dash.Columns(DashTraceNumCol).Hidden = True
End Sub

Sub Clear_TraceAddresses()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Port Continuity")
    
    ws.Range("6:" & Rows.Count).Clear ' Should match FirstDashRow in TraceAddresses macro
End Sub

