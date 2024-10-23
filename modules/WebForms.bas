Attribute VB_Name = "WebForms"
Sub FetchPrismData()
    Dim cell As Range
    Dim wb As Workbook
    Dim Dash As Worksheet
    Dim StrandQD As Worksheet
    Dim SeparateStrandQD As Boolean
    Dim FiberQD As Worksheet
    Dim HAF As Worksheet
    
    Call UnloadAllForms
    Set Dash = ThisWorkbook.Worksheets("PrismMQMS")
    Set imports = ThisWorkbook.Worksheets("File Imports")
    Dim WebForm As New PrismMQMS
    WebForm.Show vbModeless

    
    ' Load the BOM(s). Includes main BOM and optionally an extra one containing just StrandQuickDetails
    Set cell = imports.Range("Path_StrandQuickDetails")
    If Not cell.Value = Empty Then path = cell.Value
    If IsEmpty(cell) Or path = "" Or path = False Then
        Set cell = imports.Range("Path_BOMs")
        If Not cell.Value = Empty Then path = cell.Value
        If IsEmpty(cell) Or path = "" Or path = False Then
            ThisWorkbook.Sheets("File Imports").Activate
            imports.Range("Path_StrandQuickDetails").Select
            MsgBox ("Missing Path_BOMs and Path_StrandQuickDetails; Cannot fetch Fiber/Strand details for MQMS/Prism")
            WebForm.Hide
            Call UnloadAllForms
            End
        End If
    ' If the StrandQuickDetails cell isn't empty, then set StrandQD using that path
    Else
        Set wb = OpenPath(cell)
        If WorksheetExists("StrandQuickDetails", wb) Then
            Set StrandQD = wb.Worksheets("StrandQuickDetails")
            SeparateStrandQD = True
        Else
            ThisWorkbook.Sheets("File Imports").Activate
            cell.Select
            MsgBox ("Could not locate StrandQuickDetails tab; Please select another file and try again.")
            WebForm.Hide
            Call UnloadAllForms
            End
        End If
    End If
    
    ' Set FiberQD
    cell = imports.Range("Path_BOMs")
    Set wb = OpenPath(cell)
    
    If WorksheetExists("FiberQuickDetails", wb) Then
        Set FiberQD = wb.Worksheets("FiberQuickDetails")
    Else
        ThisWorkbook.Sheets("File Imports").Activate
        cell.Select
        MsgBox ("Could not locate FiberQuickDetails tab; Please select another file and try again.")
        WebForm.Hide
        Call UnloadAllForms
        End
    End If
    
    ' Set StrandQD if it isn't already
    If StrandQD Is Nothing Then
        If WorksheetExists("StrandQuickDetails", wb) Then
            Set StrandQD = wb.Worksheets("StrandQuickDetails")
            SeparateStrandQD = False
        Else
            ThisWorkbook.Sheets("File Imports").Activate
            cell.Select
            MsgBox ("Could not locate StrandQuickDetails tab; Please select another file and try again.")
            WebForm.Hide
            Call UnloadAllForms
            End
        End If
    End If
    
    
    ' Load the HAF (and a second HAF, optionally)
    Set cell = imports.Range("Path_HAF")
    Set wb = OpenPath(cell)
    Set HAF = wb.Sheets(1)
    
    Set cell = imports.Range("Path_SG_HAF")
    If Not cell.Value = Empty Then path = cell.Value
    If IsEmpty(cell) Or path = "" Or path = False Then
        Set SGHAF = Nothing
    Else
        Set wb = OpenPath(cell)
        Set SGHAF = wb.Sheets(1)
    End If
    
    ' Load the MOP
    Set cell = imports.Range("Path_MOP")
    Set wb = OpenPath(cell)
    Set MOP = wb.Worksheets("MOP")
    
    ' Update Project Details and Route Module with Strand Quick Details info
    WebForm.TotalFootage.Text = StrandQD.Range("B10").Value
    WebForm.RouteAerial.Text = StrandQD.Range("B7").Value
    WebForm.RouteUG.Text = StrandQD.Range("B8").Value + StrandQD.Range("B9").Value
    
    ' Update Project Details with the sum of addresses in the HAF
    If SGHAF Is Nothing Then
        WebForm.TotalHomesPassed.Text = HAF.Range("A" & HAF.Rows.Count).End(xlUp).row
    Else
        WebForm.TotalHomesPassed.Text = HAF.Range("A" & HAF.Rows.Count).End(xlUp).row + SGHAF.Range("A" & SGHAF.Rows.Count).End(xlUp).row
    End If
    
    ' Update MQMS Fiber fields
    WebForm.WorkspaceName.Text = "CHR_" & Dash.Range("PID").Value & "_##"
    WebForm.FiberAerial.Text = FiberQD.Range("B7").Value
    WebForm.FiberUG.Text = FiberQD.Range("B8").Value + FiberQD.Range("B9").Value
    
    With Application.WorksheetFunction
        If SGHAF Is Nothing Then
            WebForm.EstHousesAerial.Text = .CountIf(HAF.Range("O2:O257"), "AERIAL")
            WebForm.EstHousesUG.Text = .CountIf(HAF.Range("O2:O257"), "UNDERGROUND")
        Else
            WebForm.EstHousesAerial.Text = .CountIf(HAF.Range("O2:O257"), "AERIAL") + .CountIf(SGHAF.Range("O2:O257"), "AERIAL")
            WebForm.EstHousesUG.Text = .CountIf(HAF.Range("O2:O257"), "UNDERGROUND") + .CountIf(SGHAF.Range("O2:O257"), "UNDERGROUND")
        End If
    End With
    
    ' Update Prism fields
    WebForm.PrismStrandAerial.Text = StrandQD.Range("B7").Value
    WebForm.PrismStrandUG.Text = StrandQD.Range("B8").Value + StrandQD.Range("B9").Value
    WebForm.PrismFiberAerial.Text = FiberQD.Range("B7").Value
    WebForm.PrismFiberUG.Text = FiberQD.Range("B8").Value + FiberQD.Range("B9").Value
    
    Dim FoundCell As Range
    Set FoundCell = MOP.Range("B:B").Find(What:="HUB NAME")
    If Not FoundCell Is Nothing Then
        If IsEmpty(FoundCell.Offset(0, 1)) Then
            Set FoundCell = MOP.Range("B:B").FindNext(After:=FoundCell)
        End If
        WebForm.PrismCoaxHub.Text = FoundCell.Offset(0, 1).Value
        WebForm.PrismCoaxNode.Text = Dash.Range("OLT").Value & "-" & Dash.Range("OLT").Value
        WebForm.PrismHub.Text = FoundCell.Offset(0, 1).Value
        WebForm.PrismCLLI.Text = FoundCell.Offset(-1, 1).Value
        WebForm.PrismServiceGroupID.Text = "PON_" & FoundCell.Offset(-1, 1).Value & "_" & Format(Dash.Range("SUB_CIRCUIT").Value, "0000") & "-" & Format(Dash.Range("SUB_CIRCUIT").Value + 3, "0000")
    End If
    
    Set FoundCell = Nothing
    Set FoundCell = MOP.Range("I:I").Find(What:="Km Dist")
    If Not FoundCell Is Nothing Then
        WebForm.PrismEstKm.Text = Format(FoundCell.Offset(1, 0).Value, "####0.0")
        WebForm.PrismRowRack.Text = FoundCell.Offset(1, -6).Value
        WebForm.PrismSheathID.Text = FoundCell.Offset(1, -5).Value
        WebForm.PrismFiberAssigned.Text = FoundCell.Offset(1, -4).Value & "/" & FoundCell.Offset(1, -3).Value & "-" & FoundCell.Offset(2, -3).Value
    End If
    
    Set FoundCell = Nothing
    Set FoundCell = MOP.Range("C:C").Find(What:="CORWAVE")
    If Not FoundCell Is Nothing Then
        If FoundCell.Offset(1, 0).Value Like "*####.##*####.##*####.##*####.##*" Then
            WebForm.PrismWavelength.Text = Right(Split(FoundCell.Offset(1, 0).Value, ".")(0), 4) & "." & Left(Split(FoundCell.Offset(1, 0).Value, ".")(1), 2)
        End If
    End If
    
    
    ' Close extra workbooks
    FiberQD.Parent.Close
    If (Not (StrandQD Is Nothing)) And (SeparateStrandQD = True) Then StrandQD.Parent.Close
    HAF.Parent.Close
    If Not (SGHAF Is Nothing) Then SGHAF.Parent.Close
    MOP.Parent.Close
    ' Switch to the first tab in the UserForm
    WebForm.MultiPage1.Value = 0
    WebForm.Show vbModeless
End Sub
