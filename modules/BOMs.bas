Attribute VB_Name = "BOMs"
Sub CheckAllErrors_BOMs()
    Dim imports As Worksheet
    Dim Dash As Worksheet
    Dim BOMs As Workbook
    Dim BOMSheet As Worksheet
    Dim OvBOM As Workbook
    Dim Mats As Worksheet
    Dim ErrorCell As Range
    
    Call Clear_BOMDashboard
    Call Clear_ErrorDashboard("BOMs")
    
    WasBOMOpen = False
    WasOvBOMOpen = False
    Set imports = ThisWorkbook.Worksheets("File Imports")
    Set Dash = ThisWorkbook.Worksheets("BOMs")
    
    If IsWorkBookOpen(imports.[Path_BOMs]) Then WasBOMOpen = True
    Set BOMs = OpenPath(imports.[Path_BOMs])
    
    Set cell = imports.Range("Path_Overall_BOM")
    If Not cell.Value = Empty Then path = cell.Value
    If IsEmpty(cell) Or path = "" Or path = False Then
        ThisWorkbook.Sheets("File Imports").Activate
        cell.Select
        MsgBox (cell.Name.Name & " is not set. BOM analysis will skip OvBOM comparisons.")
        Set OvBOM = Nothing
        Set Mats = Nothing
    Else
        If IsWorkBookOpen(imports.[Path_Overall_BOM]) Then WasOvBOMOpen = True
        Set OvBOM = OpenPath(imports.[Path_Overall_BOM])
        Set Mats = OvBOM.Worksheets("EPON Optics and Materials")
    End If
    

    Dash.Activate 'Return focus back to the dashboard
    
'    'Might get back to this later: set up variables for all the dashboard locations, rather than using cell addresses below
'    DashBOM_TotalSheath
'    DashBOM_TotalSheathMiles
'    DashQD_AerialFBS
'    DashQD_UGFBS
'    DashQD_TotalFBS
'    DashQD_TotalSheath
'    DashQD_TotalSheathMiles
'    DashOv_AerialFBS
'    DashOv_UGFBS
'    DashOv_TotalFBS
'    DashOv_TotalSheathMiles
'    DashError_AerialFBS
'    DashError_UGFBS
'    DashError_TotalFBS
'    DashError_TotalSheath
'    DashError_TotalSheathMiles
'    DashTail_UG2CT
'    DashTail_UG4CT
'    DashTail_UG8CT
'    DashTail_UG12CT
'    DashTail_A2CT
'    DashTail_A4CT
'    DashTail_A8CT
'    DashTail_A12CT
'    DashMST_UG2CT
'    DashMST_UG4CT
'    DashMST_UG8CT
'    DashMST_UG12CT
'    DashMST_A2CT
'    DashMST_A4CT
'    DashMST_A8CT
'    DashMST_A12CT
'    DashError_UGMST2CT
'    DashError_UGMST4CT
'    DashError_UGMST8CT
'    DashError_UGMST12CT
'    DashError_AMST2CT
'    DashError_AMST4CT
'    DashError_AMST8CT
'    DashError_AMST12CT
    
    'Setting up Schema
    Dim BOMTotalSheath As SHTHSheet: Set BOMTotalSheath = New SHTHSheet
    Dim BOMInternals As INTSheet: Set BOMInternals = New INTSheet
    Dim BOMNodes As NODESheet: Set BOMNodes = New NODESheet
    Dim BOMSpliceCans As SPLSheet: Set BOMSpliceCans = New SPLSheet
    
    Set BOMTotalSheath.sheet = BOMs.Worksheets("FiberTotalSheath")
    Set BOMInternals.sheet = BOMs.Worksheets("FiberInternals")
    Set BOMNodes.sheet = BOMs.Worksheets("FiberNodes")
    Set BOMSpliceCans.sheet = BOMs.Worksheets("FiberSplices")
    
    'Setting up BOM tab column placeholders
    Dim Col_TotSheathLoc As String: Col_TotSheathLoc = BOMTotalSheath.get_letter(SHTH_LOCATION)
    Dim Col_TotSheathModel As String: Col_TotSheathModel = BOMTotalSheath.get_letter(SHTH_MODEL)
    Dim Col_TotSheathFtg As String: Col_TotSheathFtg = BOMTotalSheath.get_letter(SHTH_TOTAL_FTG)
    Dim Col_TotSheathMiles As String: Col_TotSheathMiles = BOMTotalSheath.get_letter(SHTH_TOTAL_MILES)
    Dim Col_IntModel As String: Col_IntModel = BOMInternals.get_letter(INT_MODEL)
    Dim Col_IntCount As String: Col_IntCount = BOMInternals.get_letter(INT_COUNT)
    Dim Col_NodeModel As String: Col_NodeModel = BOMNodes.get_letter(NODE_MODEL)
    Dim Col_NodeCount As String: Col_NodeCount = BOMNodes.get_letter(NODE_COUNT)
    Dim Col_SplModel As String: Col_SplModel = BOMSpliceCans.get_letter(SPL_MODEL)
    Dim Col_SplCount As String: Col_SplCount = BOMSpliceCans.get_letter(SPL_COUNT)
    
    'Setting up FiberQuickDetails placeholders
    
    '--Comparing QuickDetails to TotalSheath--
    Set BOMSheet = BOMTotalSheath.sheet
    Dash.Range("C9") = Application.WorksheetFunction.Sum(BOMSheet.Range(Col_TotSheathFtg & ":" & Col_TotSheathFtg))
    Dash.Range("C10") = Application.WorksheetFunction.Sum(BOMSheet.Range(Col_TotSheathMiles & ":" & Col_TotSheathMiles))
    
    Set BOMSheet = BOMs.Worksheets("FiberQuickDetails")
    Dash.Range("D6") = BOMSheet.Range("B7")
    Dash.Range("D7") = BOMSheet.Range("B8").Value + BOMSheet.Range("B9").Value
    Dash.Range("D8") = BOMSheet.Range("B11")
    Dash.Range("D9") = BOMSheet.Range("B15")
    Dash.Range("D10") = BOMSheet.Range("C15")
    
    If OvBOM Is Nothing Then
        Dash.Range("E6:E8").Value = "X"
        Dash.Range("E10").Value = "X"
    Else
        
        
        
        
        Dash.Range("E6") = OvBOM.Worksheets("Project Overview").Range("F12")
        Dash.Range("E7") = OvBOM.Worksheets("Project Overview").Range("F13")
        Dash.Range("E8") = Dash.Range("E6").Value + Dash.Range("E7").Value
        Dash.Range("E10") = OvBOM.Worksheets("Project Overview").Range("F9")
    End If
    
    
    '--Comparing tail footage to MSTs--
    Set BOMSheet = BOMTotalSheath.sheet
    OLT_Tail_UG = IIf(Dash.Range("OLT_Tail_Location").Value = "Underground", Dash.Range("OLT_Tail_Count") * 100, 0)
    OLT_Tail_Aerial = IIf(Dash.Range("OLT_Tail_Location").Value = "Aerial", Dash.Range("OLT_Tail_Count") * 100, 0)
    Dash.Range("C13") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_tail_2ct, "U")
    Dash.Range("C14") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_tail_4ct, "U")
    Dash.Range("C15") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_tail_8ct, "U")
    Dash.Range("C16") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_tail_12ct, "U") - OLT_Tail_UG
    Dash.Range("C17") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_tail_2ct, "A")
    Dash.Range("C18") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_tail_4ct, "A")
    Dash.Range("C19") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_tail_8ct, "A")
    Dash.Range("C20") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_tail_12ct, "A") - OLT_Tail_Aerial
    
    Set BOMSheet = BOMNodes.sheet
    Dash.Range("E13") = SumBOMModels(BOMSheet, Col_NodeModel, Col_NodeCount, mst_2ct, "U")
    Dash.Range("E14") = SumBOMModels(BOMSheet, Col_NodeModel, Col_NodeCount, mst_4ct, "U")
    Dash.Range("E15") = SumBOMModels(BOMSheet, Col_NodeModel, Col_NodeCount, mst_8ct, "U")
    Dash.Range("E16") = SumBOMModels(BOMSheet, Col_NodeModel, Col_NodeCount, mst_12ct, "U")
    Dash.Range("E17") = SumBOMModels(BOMSheet, Col_NodeModel, Col_NodeCount, mst_2ct, "A")
    Dash.Range("E18") = SumBOMModels(BOMSheet, Col_NodeModel, Col_NodeCount, mst_4ct, "A")
    Dash.Range("E19") = SumBOMModels(BOMSheet, Col_NodeModel, Col_NodeCount, mst_8ct, "A")
    Dash.Range("E20") = SumBOMModels(BOMSheet, Col_NodeModel, Col_NodeCount, mst_12ct, "A")
    
    '--Comparing BOMs to Overall BOM--
    
    'OLT
    Set BOMSheet = BOMs.Worksheets("FiberCabinets")
    Dash.Range("D23") = SumBOMModels(BOMSheet, "D", "E", "FTTX_VN_ENTRA_SF-4X_OLT")
    Dash.Range("E23") = SumOvBOMModels(Mats, "Vecima", "NODE 4P 10G EPON R-OLT 4/XFP 4/Port Lic")
    
    'Splitters
    Set BOMSheet = BOMInternals.sheet
    Dash.Range("D25") = SumBOMModels(BOMSheet, Col_IntModel, Col_IntCount, "FTTX_CO_1X2_SPL")
    Dash.Range("D26") = SumBOMModels(BOMSheet, Col_IntModel, Col_IntCount, "FTTX_CO_1X32_SPL")
    Dash.Range("D27") = SumBOMModels(BOMSheet, Col_IntModel, Col_IntCount, "FTTX_CO_1X64_SPL")
    Dash.Range("E25") = SumOvBOMModels(Mats, "Commscope", "FIBER SPLITTER UNIV 2 WAY FBT BARE 1M INPUT 1M LOOSE OUTPUT")
    Dash.Range("E26") = SumOvBOMModels(Mats, "Commscope", "1x32 Field Install Splitter with carrier")
    Dash.Range("E27") = SumOvBOMModels(Mats, "Commscope", "1x64 Field Install Splitter with carrier")
    
    'OTEs
    Set BOMSheet = BOMNodes.sheet
    Dash.Range("D29") = SumBOMModels(BOMSheet, Col_NodeModel, Col_NodeCount, ote_2ct)
    Dash.Range("D30") = SumBOMModels(BOMSheet, Col_NodeModel, Col_NodeCount, ote_4ct)
    Dash.Range("D31") = SumBOMModels(BOMSheet, Col_NodeModel, Col_NodeCount, ote_8ct)
    Dash.Range("D32") = SumBOMModels(BOMSheet, Col_NodeModel, Col_NodeCount, ote_12ct)
    Dash.Range("D33") = Application.WorksheetFunction.Sum(Dash.Range("D29:D32"))
    Dash.Range("E29") = SumOvBOMModels(Mats, "Commscope", "FTTX,OTE-MINI, 2 Port, No Splitter, w/Ground")
    Dash.Range("E30") = SumOvBOMModels(Mats, "Commscope", "FTTX,OTE-MINI, 4 Port, No Splitter, w/Ground")
    Dash.Range("E31") = SumOvBOMModels(Mats, "Commscope", "FTTX,OTE-MINI, 8 Port, No Splitter, w/Ground")
    Dash.Range("E32") = SumOvBOMModels(Mats, "Commscope", "FTTX,OTE-MINI, 12 Port, No Splitter, w/Ground")
    Dash.Range("E33") = SumOvBOMModels(Mats, "Commscope", "Bracket, OTE-MINI Strand Hanger")
    
    'MSTs
    'Counts for UG/Aerial were found separately earlier, so now we can sum them
    Dash.Range("D35") = Dash.Range("E13") + Dash.Range("E17")
    Dash.Range("D36") = Dash.Range("E14") + Dash.Range("E18")
    Dash.Range("D37") = Dash.Range("E15") + Dash.Range("E19")
    Dash.Range("D38") = Dash.Range("E16") + Dash.Range("E20")
    Dash.Range("E35") = SumOvBOMModels(Mats, "Commscope", "MST 2-Port with 100ft Toneable Drop")
    Dash.Range("E36") = SumOvBOMModels(Mats, "Commscope", "MST 4-Port with 100ft Toneable Drop")
    Dash.Range("E37") = SumOvBOMModels(Mats, "Commscope", "MST 8-Port with 100ft Toneable Drop")
    Dash.Range("E38") = SumOvBOMModels(Mats, "Commscope", "MST 12-Port with 100ft Toneable Drop") 'There is no entry for this in the OvBOM
    
    'Splice Enclosures
    Set BOMSheet = BOMSpliceCans.sheet
    Dash.Range("D40") = SumBOMModels(BOMSheet, "D", "E", "*B66_SPL_DIST")
    Dash.Range("D41") = SumBOMModels(BOMSheet, "D", "E", "*B66_SPL_SPLIT")
    Dash.Range("D42") = SumBOMModels(BOMSheet, "D", "E", "*B66_SPL_TRUNK")
    Dash.Range("D43") = Dash.Range("D40") + Dash.Range("D41") + Dash.Range("D42")
    'Dash.Range("E40") = "N/A"
    'Dash.Range("E41") = "N/A"
    'Dash.Range("E42") = "N/A"
    Dash.Range("E43") = SumOvBOMModels(Mats, "Commscope", "FOSC B6 Closure, 1(B) Splice Tray, 3 Feed Thru Lugs and 6 Cable Ports")
    
    'Sheaths
    Set BOMSheet = BOMTotalSheath.sheet
    Dash.Range("C46") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_sheath_6ct)
    Dash.Range("C47") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_sheath_12ct) + IIf(Dash.Range("OLT_Tail_On_OvBOM").Value = "Yes", Dash.Range("OLT_Tail_Count") * 100, 0)
    Dash.Range("C48") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_sheath_24ct)
    Dash.Range("C49") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_sheath_48ct)
    Dash.Range("C50") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_sheath_72ct)
    Dash.Range("C51") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_sheath_96ct)
    Dash.Range("C52") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_sheath_144ct)
    Dash.Range("C53") = SumBOMModels(BOMSheet, Col_TotSheathModel, Col_TotSheathFtg, fiber_sheath_288ct)
    
    For Each cell In Dash.Range("C46:C53").Cells
        cell.Offset(0, 1).Value = Round(cell.Value * 1.13, 0)
    Next cell
    
    Dash.Range("E46") = SumOvBOMModels(Mats, "Commscope", "FBR 6 CT ARMORED LT SM DRY")
    Dash.Range("E47") = SumOvBOMModels(Mats, "Commscope", "FBR 12 CT ARMORED LT SM DRY")
    Dash.Range("E48") = SumOvBOMModels(Mats, "Commscope", "Fiber 24CT Single Armor Loose Tube Gel Free")
    Dash.Range("E49") = SumOvBOMModels(Mats, "Commscope", "Fiber 48CT Single Armor Loose Tube Gel Free")
    Dash.Range("E50") = SumOvBOMModels(Mats, "Commscope", "Fiber 72CT Single Armor Loose Tube Gel Free")
    Dash.Range("E51") = SumOvBOMModels(Mats, "Commscope", "Fiber 96CT Single Armor Loose Tube Gel Free")
    Dash.Range("E52") = SumOvBOMModels(Mats, "Commscope", "Fiber 144CT Single Armor Loose Tube Gel Free")
    Dash.Range("E53") = SumOvBOMModels(Mats, "Commscope", "Fiber 288CT Single Armor Loose Tube Gel Free")
    
    '--Checking for Errors--
    
    For Each cell In Dash.Range("F6:F10")
        ' This line is required to avoid a argument type mismatch (cell iterator is not technically a "range")
        Set ErrorCell = cell
        
        SheathBOMVal = ErrorCell.Offset(0, -3).Value
        QuickDetailsVal = ErrorCell.Offset(0, -2).Value
        OvBOMVal = ErrorCell.Offset(0, -1).Value
        If (Not Val(SheathBOMVal) = Val(QuickDetailsVal)) And (Not SheathBOMVal = "---") Then
            Call AddError("Error_BOM_TotalSheathVsQuickDetails", ErrorCell.Offset(0, -4).Value, "FiberTotalSheath doesn't match QuickDetails", ErrorCell)
            ErrorCell.Font.Color = vbRed
        End If
        
        If OvBOMVal = "X" Then
            ErrorCell.Offset(0, -1).HorizontalAlignment = xlCenter
        ElseIf (Not Val(QuickDetailsVal) = Val(OvBOMVal)) And (Not OvBOMVal = "---") Then
            Call AddError("Error_BOM_OverallVsQuickDetails", ErrorCell.Offset(0, -4).Value, "Overall BOM doesn't match BOMs", ErrorCell)
            ErrorCell.Font.Color = vbRed
        End If
    Next cell
    
    Dim thrower As String
    For Each cell In Dash.Range("F13:F20")
        Set ErrorCell = cell
        
        If Not (ErrorCell.Offset(0, -3).Value = ErrorCell.Offset(0, -1).Value * 100) Then
            thrower = ErrorCell.Offset(0, -4).Value
            thrower = Right(thrower, Len(thrower) - InStrRev(thrower, "_"))
            thrower = thrower & " count tails / MSTs (" & ErrorCell.Offset(0, -5).Value & ")"
            Call AddError("Error_BOM_TailsVsMSTs", thrower, "Tail footage doesn't match MST count", ErrorCell)
            ErrorCell.Font.Color = vbRed
        End If
    Next cell
    
    For Each cell In Dash.Range("F23:F43")
        Set ErrorCell = cell
        
        If ErrorCell.Offset(0, -1).Value = "X" Then
            ErrorCell.Offset(0, -1).HorizontalAlignment = xlCenter
        ElseIf (Not ErrorCell.Offset(0, -2) = ErrorCell.Offset(0, -1)) And (Not ErrorCell.Offset(0, -1) = "---") Then
            Call AddError("Error_BOM_OverallVsBOMs", ErrorCell.Offset(0, -4).Value, "Overall BOM doesn't match BOMs", ErrorCell)
            ErrorCell.Font.Color = vbRed
        End If
    Next cell
    
    For Each cell In Dash.Range("F46:F53")
        Set ErrorCell = cell
        
        If ErrorCell.Offset(0, -1).Value = "X" Then
            ErrorCell.Offset(0, -1).HorizontalAlignment = xlCenter
        Else
            diff = Abs(ErrorCell.Offset(0, -2).Value - ErrorCell.Offset(0, -1).Value)
            If diff > 1 Then
                Call AddError("Error_BOM_OverallVsBOMs", ErrorCell.Offset(0, -4).Value, "Overall BOM doesn't match BOMs", ErrorCell)
                ErrorCell.Font.Color = vbRed
            End If
        End If
    Next cell
    
    Call CheckForMissedModels(BOMNodes.sheet, Col_NodeModel, Col_NodeCount, all_hybrids)
    Call CheckForMissedModels(BOMInternals.sheet, Col_IntModel, Col_IntCount, all_internals)
    Call CheckForMissedModels(BOMTotalSheath.sheet, Col_TotSheathModel, Col_TotSheathFtg, all_fibers, Col_TotSheathLoc)
    If Dash.Range("B56").Value = "" Then
        Dash.Range("B56").Value = "No Misc Errors"
        Dash.Range("B56").Font.Color = vbGreen
    End If
    
    If WasBOMOpen = False Then BOMs.Close
    If (Not OvBOM Is Nothing) And (WasOvBOMOpen = False) Then OvBOM.Close
End Sub

Function SumBOMModels(BOMTab As Worksheet, SearchCol As String, CountCol As String, model As Variant, Optional Loc As String = "") As Variant
    Dim LR As Integer
    Dim SearchRange As Range
    Dim Sum As Variant: Sum = 0
    
    'Find the last row of the targeted sheet
    LR = BOMTab.Range(CountCol & BOMTab.Rows.Count).End(xlUp).row
    'Build the range the search will loop through
    Set SearchRange = BOMTab.Range(SearchCol & "2:" & SearchCol & LR)
    
    
    If TypeOf model Is Dictionary Then
        'If we're looking for a dictionary, we'll check whether the model in each row matches an entry in the dictionary
        For Each cell In SearchRange
            If model.Exists(cell.Value) Then
                If Loc = "" Then
                    Sum = Sum + BOMTab.Range(CountCol & cell.row)
                Else
                    If model(cell.Value) = Loc Then
                        Sum = Sum + BOMTab.Range(CountCol & cell.row)
                    End If
                End If
            End If
        Next cell
    ElseIf VarType(model) = vbString Then
        'If we're looking for a string, we'll loop through the sheet once and check for a match.
        For Each cell In SearchRange
            'If the model is a match, add the Count (or Footage) column value to the running total
            If cell.Value Like model Then
                Sum = Sum + BOMTab.Range(CountCol & cell.row)
            End If
        Next cell
    Else
        MsgBox "Invalid call to SumBOMModels, alert the maintainer of this spreadsheet to the issue."
        End
    End If
    
    SumBOMModels = Sum
End Function

Function SumOvBOMModels(OvBOM As Worksheet, supplier As String, desc As String) As Variant
    If OvBOM Is Nothing Then
         SumOvBOMModels = "X"
    Else
        Dim Sum As Variant
        Sum = Application.WorksheetFunction.SumIfs(OvBOM.Range("F1:F300"), OvBOM.Range("A1:A300"), supplier, OvBOM.Range("D1:D300"), desc)
        SumOvBOMModels = Sum
    End If
End Function

Function CheckForMissedModels(BOMTab As Worksheet, SearchCol As String, CountCol As String, modelDict As Dictionary, Optional LocCol As String = "")
    Dim LR As Integer
    Dim SearchRange As Range
    Dim ErrorCell As Range
    
    'Find the last row of the targeted sheet
    LR = BOMTab.Range(SearchCol & BOMTab.Rows.Count).End(xlUp).row
    'Build the range the search will loop through
    Set SearchRange = BOMTab.Range(SearchCol & "2:" & SearchCol & LR)
    'Define the first error cell
    With ThisWorkbook.Worksheets("BOMs")
        Set ErrorCell = .Range("B" & .Rows.Count).End(xlUp).Offset(1, 0)
    End With

    
    For Each cell In SearchRange
        If modelDict.Exists(cell.Value) Then
            If Not LocCol = "" Then
                If Not BOMTab.Range(LocCol & cell.row).Value = modelDict(cell.Value) Then
                    If modelDict(cell.Value) = "A" Then
                        Call AddError("Error_BOM_HybridLocMismatch", cell.Value & "at non-Aerial location", "Aerial equipment '" & cell.Value & "' located at non-Aerial location in BOM tab " & BOMTab.Name, ErrorCell)
                        ErrorCell.Font.Color = vbRed
                        Set ErrorCell = ErrorCell.Offset(1, 0)
                    ElseIf modelDict(cell.Value) = "U" Then
                        Call AddError("Error_BOM_HybridLocMismatch", cell.Value & "at non-UG location", "Underground equipment '" & cell.Value & "' located at non-UG location in BOM tab " & BOMTab.Name, ErrorCell)
                        ErrorCell.Font.Color = vbRed
                        Set ErrorCell = ErrorCell.Offset(1, 0)
                    End If
                End If
            End If
        Else
            Call AddError("Error_BOM_UnknownModel", cell.Value, "Can't recognize model in tab '" & BOMTab.Name & "': " & cell.Value & " (Count: " & BOMTab.Range(CountCol & cell.row).Value & ")", ErrorCell, True)
            ErrorCell.Font.ColorIndex = 44
            Set ErrorCell = ErrorCell.Offset(1, 0)
        End If
    Next cell
End Function

Sub Clear_BOMDashboard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("BOMs")
    
    ws.Range("C6:F10").Clear
    ws.Range("C6:C8").Value = "'---"
    ws.Range("C6:C8").HorizontalAlignment = xlCenter
    ws.Range("E9").Value = "'---"
    ws.Range("E9").HorizontalAlignment = xlCenter
    
    ws.Range("C13:C20").Clear
    ws.Range("E13:F20").Clear
    
    ws.Range("D23:F43").Clear
    ws.Range("E40:E42").Value = "'---"
    ws.Range("E40:E42").HorizontalAlignment = xlCenter
    
    ws.Range("C46:F53").Clear
    
    ws.Range("B56:B" & ws.Rows.Count).Clear
    ws.Range("B56:B" & ws.Rows.Count).ClearFormats
End Sub
