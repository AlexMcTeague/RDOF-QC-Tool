Attribute VB_Name = "BOM_CONFIG"
Private FQD_LR_p As Integer
Private FQD_FBS_Row_p As Integer
Private FQD_FBS_Aerial_p As Range
Private FQD_FBS_Aerial_Miles_p As Range
Private FQD_FBS_UG_p As Range
Private FQD_FBS_UG_Miles_p As Range
Private FQD_FBS_Riser_p As Range
Private FQD_FBS_Riser_Miles_p As Range
Private FQD_Total_Sheath_p As Range
Private FQD_Total_Sheath_Miles_p As Range

Private all_hybrids_p As Dictionary
Private all_internals_p As Dictionary
Private all_fibers_p As Dictionary

Private mst_2ct_p As Dictionary
Private mst_4ct_p As Dictionary
Private mst_8ct_p As Dictionary
Private mst_12ct_p As Dictionary
Private ote_2ct_p As Dictionary
Private ote_4ct_p As Dictionary
Private ote_8ct_p As Dictionary
Private ote_12ct_p As Dictionary
Private splice_can_p As Dictionary
Private splitter_p As Dictionary
Private misc_internal_p As Dictionary
Private fiber_tail_2ct_p As Dictionary
Private fiber_tail_4ct_p As Dictionary
Private fiber_tail_8ct_p As Dictionary
Private fiber_tail_12ct_p As Dictionary
Private fiber_sheath_6ct_p As Dictionary
Private fiber_sheath_12ct_p As Dictionary
Private fiber_sheath_24ct_p As Dictionary
Private fiber_sheath_48ct_p As Dictionary
Private fiber_sheath_72ct_p As Dictionary
Private fiber_sheath_96ct_p As Dictionary
Private fiber_sheath_144ct_p As Dictionary
Private fiber_sheath_288ct_p As Dictionary

Public Property Get FQD_LR(sheet As Worksheet) As Integer
    If FQD_LR_p Is Nothing Then
        LRA = sheet.Range("A" & sheet.Rows.Count).End(xlUp).row
        LRB = sheet.Range("B" & sheet.Rows.Count).End(xlUp).row
        FQD_LR_p = WorksheetFunction.Max(LRA, LRB)
    End If
    FQD_LR = FQD_LR_p
End Property

Public Property Get FQD_FBS_Row(sheet As Worksheet) As Integer
    If FQD_FBS_Row_p Is Nothing Then
        For i = 1 To FQD_LR
            If sheet.Cells(i, "A").Value = "Fiber Bearing Strand" Then
                FQD_FBS_Row_p = i
                Exit For
            End If
        Next i
    End If
    FQD_FBS_Row = FQD_FBS_Row_p
End Property

Public Property Get FQD_Totals_Row(sheet As Worksheet) As Integer
    If FQD_Totals_Row_p Is Nothing Then
        For i = 1 To FQD_LR
            If sheet.Cells(i, "A").Value = "Totals" Then
                FQD_Totals_Row_p = i
                Exit For
            End If
        Next i
    End If
    FQD_Totals_Row = FQD_Totals_Row_p
End Property

Public Property Get FQD_FBS_Aerial(sheet As Worksheet) As Range
    If FQD_FBS_Aerial_p Is Nothing Then
        For i = FQD_FBS_Row To FQD_LR
            If sheet.Cells(i, "A").Value = "Aerial" Then
                Set FQD_FBS_Aerial_p = sheet.Range("B" & i)
                Exit For
            End If
        Next i
    End If
    Set FQD_FBS_Aerial = FQD_FBS_Aerial_p
End Property

Public Property Get FQD_FBS_Aerial_Miles(sheet As Worksheet) As Range
    If FQD_FBS_Aerial_Miles_p Is Nothing Then
        For i = FQD_FBS_Row To FQD_LR
            If sheet.Cells(i, "A").Value = "Aerial" Then
                Set FQD_FBS_Aerial_Miles_p = sheet.Range("C" & i)
                Exit For
            End If
        Next i
    End If
    Set FQD_FBS_Aerial_Miles = FQD_FBS_Aerial_Miles_p
End Property

Public Property Get FQD_FBS_UG(sheet As Worksheet) As Range
    If FQD_FBS_UG_p Is Nothing Then
        For i = FQD_FBS_Row To FQD_LR
            If sheet.Cells(i, "A").Value = "Underground" Then
                Set FQD_FBS_UG_p = sheet.Range("B" & i)
                Exit For
            End If
        Next i
    End If
    Set FQD_FBS_UG = FQD_FBS_UG_p
End Property

Public Property Get FQD_FBS_UG_Miles(sheet As Worksheet) As Range
    If FQD_FBS_UG_Miles_p Is Nothing Then
        For i = FQD_FBS_Row To FQD_LR
            If sheet.Cells(i, "A").Value = "Underground" Then
                Set FQD_FBS_UG_Miles_p = sheet.Range("C" & i)
                Exit For
            End If
        Next i
    End If
    Set FQD_FBS_UG_Miles = FQD_FBS_UG_Miles_p
End Property

Public Property Get FQD_FBS_Riser(sheet As Worksheet) As Range
    If FQD_FBS_Riser_p Is Nothing Then
        For i = FQD_FBS_Row To FQD_LR
            If sheet.Cells(i, "A").Value = "Riser" Then
                Set FQD_FBS_Riser_p = sheet.Range("B" & i)
                Exit For
            End If
        Next i
    End If
    Set FQD_FBS_Riser = FQD_FBS_Riser_p
End Property

Public Property Get FQD_FBS_Riser_Miles(sheet As Worksheet) As Range
    If FQD_FBS_Riser_Miles_p Is Nothing Then
        For i = FQD_FBS_Row To FQD_LR
            If sheet.Cells(i, "A").Value = "Riser" Then
                Set FQD_FBS_Riser_Miles_p = sheet.Range("C" & i)
                Exit For
            End If
        Next i
    End If
    Set FQD_FBS_Riser_Miles = FQD_FBS_Riser_Miles_p
End Property

Public Property Get FQD_Total_Sheath(sheet As Worksheet) As Range
    If FQD_Total_Sheath_p Is Nothing Then
        For i = FQD_FBS_Row To FQD_LR
            If sheet.Cells(i, "A").Value = "Sheath" Then
                Set FQD_Total_Sheath_p = sheet.Range("B" & i)
                Exit For
            End If
        Next i
    End If
    Set FQD_Total_Sheath = FQD_Total_Sheath_p
End Property

Public Property Get FQD_Total_Sheath_Miles(sheet As Worksheet) As Range
    If FQD_Total_Sheath_Miles_p Is Nothing Then
        For i = FQD_FBS_Row To FQD_LR
            If sheet.Cells(i, "A").Value = "Sheath" Then
                Set FQD_Total_Sheath_Miles_p = sheet.Range("C" & i)
                Exit For
            End If
        Next i
    End If
    Set FQD_Total_Sheath_Miles = FQD_Total_Sheath_Miles_p
End Property



' All Hybrids
Public Property Get all_hybrids() As Dictionary
    If all_hybrids_p Is Nothing Then
        Set all_hybrids_p = New Dictionary
        Set all_hybrids_p = MergeDicts(mst_2ct, mst_4ct, mst_8ct, mst_12ct, ote_2ct, ote_4ct, ote_8ct, ote_12ct)
    End If
    Set all_hybrids = all_hybrids_p
End Property

'All Internals
Public Property Get all_internals() As Dictionary
    If all_internals_p Is Nothing Then
        Set all_internals_p = New Dictionary
        Set all_internals_p = MergeDicts(splitter, misc_internal)
    End If
    Set all_internals = all_internals_p
End Property

'All Fiber Sheaths & Tails
Public Property Get all_fibers() As Dictionary
    If all_fibers_p Is Nothing Then
        Set all_fibers_p = New Dictionary
        Set all_fibers_p = MergeDicts(fiber_tail_2ct, fiber_tail_4ct, fiber_tail_8ct, fiber_tail_12ct, fiber_sheath_6ct, fiber_sheath_12ct, fiber_sheath_24ct, fiber_sheath_48ct, fiber_sheath_72ct, fiber_sheath_96ct, fiber_sheath_144ct, fiber_sheath_288ct)
    End If
    Set all_fibers = all_fibers_p
End Property



' 2CT MSTs
Public Property Get mst_2ct() As Dictionary
    If mst_2ct_p Is Nothing Then
        Set mst_2ct_p = New Dictionary
        With mst_2ct_p
            .add "FTTX_CL_U_02_SMST", "U"
            .add "FTTX_CO_U_02_NHM_SMST", "U"
            .add "FTTX_CO_A_02_MH_HMST", "A"
        End With
    End If
    Set mst_2ct = mst_2ct_p
End Property

' 4CT MSTs
Public Property Get mst_4ct() As Dictionary
    If mst_4ct_p Is Nothing Then
        Set mst_4ct_p = New Dictionary
        With mst_4ct_p
            .add "FTTX_CO_U_04_NHM_SMST", "U"
            .add "FTTX_CO_A_04_MH_HMST", "A"
        End With
    End If
    Set mst_4ct = mst_4ct_p
End Property

' 8CT MSTs
Public Property Get mst_8ct() As Dictionary
    If mst_8ct_p Is Nothing Then
        Set mst_8ct_p = New Dictionary
        With mst_8ct_p
            .add "FTTX_CO_U_08_NHM_SMST", "U"
            .add "FTTX_CO_A_08_MH_HMST", "A"
        End With
    End If
    Set mst_8ct = mst_8ct_p
End Property

' 12CT MSTs
Public Property Get mst_12ct() As Dictionary
    If mst_12ct_p Is Nothing Then
        Set mst_12ct_p = New Dictionary
        With mst_12ct_p
            .add "FTTX_CO_U_12_NHM_SMST", "U"
            .add "FTTX_CL_U_12_SMST", "U"
            .add "FTTX_CO_A_12_MH_HMST", "A"
        End With
    End If
    Set mst_12ct = mst_12ct_p
End Property

' 2CT OTEs
Public Property Get ote_2ct() As Dictionary
    If ote_2ct_p Is Nothing Then
        Set ote_2ct_p = New Dictionary
        With ote_2ct_p
            .add "FTTX_U_OTE_02", "U"
            .add "FTTX_CO_U_02_OTE", "U"
            .add "FTTX_A_OTE_02", "A"
            .add "FTTX_CO_A_02_OTE", "A"
        End With
    End If
    Set ote_2ct = ote_2ct_p
End Property

' 4CT OTEs
Public Property Get ote_4ct() As Dictionary
    If ote_4ct_p Is Nothing Then
        Set ote_4ct_p = New Dictionary
        With ote_4ct_p
            .add "FTTX_U_OTE_04", "U"
            .add "FTTX_CO_U_04_OTE", "U"
            .add "FTTX_A_OTE_04", "A"
            .add "FTTX_CO_A_04_OTE", "A"
        End With
    End If
    Set ote_4ct = ote_4ct_p
End Property

' 8CT OTEs
Public Property Get ote_8ct() As Dictionary
    If ote_8ct_p Is Nothing Then
        Set ote_8ct_p = New Dictionary
        With ote_8ct_p
            .add "FTTX_U_OTE_08", "U"
            .add "FTTX_CO_U_08_OTE", "U"
            .add "FTTX_A_OTE_08", "A"
            .add "FTTX_CO_A_08_OTE", "A"
        End With
    End If
    Set ote_8ct = ote_8ct_p
End Property

' 12CT OTEs
Public Property Get ote_12ct() As Dictionary
    If ote_12ct_p Is Nothing Then
        Set ote_12ct_p = New Dictionary
        With ote_12ct_p
            .add "FTTX_U_OTE_12", "U"
            .add "FTTX_CO_U_12_OTE", "U"
            .add "FTTX_A_OTE_12", "A"
            .add "FTTX_CO_A_12_OTE", "A"
        End With
    End If
    Set ote_12ct = ote_12ct_p
End Property

' Splice Cans
Public Property Get splice_can() As Dictionary
    If splice_can_p Is Nothing Then
        Set splice_can_p = New Dictionary
        With splice_can_p
            .add "FTTX_CO_450B66_SPL_DIST", ""
            .add "FTTX_CO_450DD66_SPL_DIST", ""
            .add "FTTX_SPL_DIST", ""
            .add "FTTX_CO_450B66_SPL_SPLIT", ""
            .add "FTTX_CO_450DD66_SPL_SPLIT", ""
            .add "FTTX_SPL_SPLIT", ""
            .add "FTTX_CO_450B66_SPL_TRUNK", ""
            .add "FTTX_CO_450DD66_SPL_TRUNK", ""
            .add "FTTX_SPL_TRUNK", ""
        End With
    End If
    Set splice_can = splice_can_p
End Property

' Splitters
Public Property Get splitter() As Dictionary
    If splitter_p Is Nothing Then
        Set splitter_p = New Dictionary
        With splitter_p
            .add "FTTX_CO_1X2_SPL", ""
            .add "FTTX_CO_1X32_SPL", ""
            .add "FTTX_CO_1X64_SPL", ""
        End With
    End If
    Set splitter = splitter_p
End Property

' Misc Internals
Public Property Get misc_internal() As Dictionary
    If misc_internal_p Is Nothing Then
        Set misc_internal_p = New Dictionary
        With misc_internal_p
            .add "DWDM_*", ""       'TODO: handle entries that use a wildcard (asterisk) differently
            .add "*MUX*", ""        'Didn't have time to implement this feature
        End With
    End If
    Set misc_internal = misc_internal_p
End Property

' 2CT Tails
Public Property Get fiber_tail_2ct() As Dictionary
    If fiber_tail_2ct_p Is Nothing Then
        Set fiber_tail_2ct_p = New Dictionary
        With fiber_tail_2ct_p
            .add "FTTX_TAIL_U_FTTX_CG_024CT_LS_2", "U"
            .add "FTTX_TAIL_U_FTTX_CG_048CT_LS_2", "U"
            .add "FTTX_TAIL_U_FTTX_CG_144CT_LS_2", "U"
            .add "FTTX_TAIL_U_FTTX_CG_288CT_LS_2", "U"
            .add "FTTX_TAIL_A_FTTX_CG_024CT_LS_2", "A"
            .add "FTTX_TAIL_A_FTTX_CG_048CT_LS_2", "A"
            .add "FTTX_TAIL_A_FTTX_CG_144CT_LS_2", "A"
            .add "FTTX_TAIL_A_FTTX_CG_288CT_LS_2", "A"
            .add "FTTX_TAIL_U_CG_024 CT_LS_2", "U"
            .add "FTTX_TAIL_U_CG_048 CT_LS_2", "U"
            .add "FTTX_TAIL_U_CG_144 CT_LS_2", "U"
            .add "FTTX_TAIL_U_CG_288 CT_LS_2", "U"
            .add "FTTX_TAIL_A_CG_024 CT_LS_2", "A"
            .add "FTTX_TAIL_A_CG_048 CT_LS_2", "A"
            .add "FTTX_TAIL_A_CG_144 CT_LS_2", "A"
            .add "FTTX_TAIL_A_CG_288 CT_LS_2", "A"
            .add "FTTX_TAIL_U_024 CT_2", "U"
            .add "FTTX_TAIL_U_048 CT_2", "U"
            .add "FTTX_TAIL_U_144 CT_2", "U"
            .add "FTTX_TAIL_U_288 CT_2", "U"
            .add "FTTX_TAIL_A_024 CT_2", "A"
            .add "FTTX_TAIL_A_048 CT_2", "A"
            .add "FTTX_TAIL_A_144 CT_2", "A"
            .add "FTTX_TAIL_A_288 CT_2", "A"
        End With
    End If
    Set fiber_tail_2ct = fiber_tail_2ct_p
End Property

' 4CT Tails
Public Property Get fiber_tail_4ct() As Dictionary
    If fiber_tail_4ct_p Is Nothing Then
        Set fiber_tail_4ct_p = New Dictionary
        With fiber_tail_4ct_p
            .add "FTTX_TAIL_U_FTTX_CG_024CT_LS_4", "U"
            .add "FTTX_TAIL_U_FTTX_CG_048CT_LS_4", "U"
            .add "FTTX_TAIL_U_FTTX_CG_144CT_LS_4", "U"
            .add "FTTX_TAIL_U_FTTX_CG_288CT_LS_4", "U"
            .add "FTTX_TAIL_A_FTTX_CG_024CT_LS_4", "A"
            .add "FTTX_TAIL_A_FTTX_CG_048CT_LS_4", "A"
            .add "FTTX_TAIL_A_FTTX_CG_144CT_LS_4", "A"
            .add "FTTX_TAIL_A_FTTX_CG_288CT_LS_4", "A"
            .add "FTTX_TAIL_U_CG_024 CT_LS_4", "U"
            .add "FTTX_TAIL_U_CG_048 CT_LS_4", "U"
            .add "FTTX_TAIL_U_CG_144 CT_LS_4", "U"
            .add "FTTX_TAIL_U_CG_288 CT_LS_4", "U"
            .add "FTTX_TAIL_A_CG_024 CT_LS_4", "A"
            .add "FTTX_TAIL_A_CG_048 CT_LS_4", "A"
            .add "FTTX_TAIL_A_CG_144 CT_LS_4", "A"
            .add "FTTX_TAIL_A_CG_288 CT_LS_4", "A"
            .add "FTTX_TAIL_U_024 CT_4", "U"
            .add "FTTX_TAIL_U_048 CT_4", "U"
            .add "FTTX_TAIL_U_144 CT_4", "U"
            .add "FTTX_TAIL_U_288 CT_4", "U"
            .add "FTTX_TAIL_A_024 CT_4", "A"
            .add "FTTX_TAIL_A_048 CT_4", "A"
            .add "FTTX_TAIL_A_144 CT_4", "A"
            .add "FTTX_TAIL_A_288 CT_4", "A"
        End With
    End If
    Set fiber_tail_4ct = fiber_tail_4ct_p
End Property

' 8CT Tails
Public Property Get fiber_tail_8ct() As Dictionary
    If fiber_tail_8ct_p Is Nothing Then
        Set fiber_tail_8ct_p = New Dictionary
        With fiber_tail_8ct_p
            .add "FTTX_TAIL_U_FTTX_CG_024CT_LS_8", "U"
            .add "FTTX_TAIL_U_FTTX_CG_048CT_LS_8", "U"
            .add "FTTX_TAIL_U_FTTX_CG_144CT_LS_8", "U"
            .add "FTTX_TAIL_U_FTTX_CG_288CT_LS_8", "U"
            .add "FTTX_TAIL_A_FTTX_CG_024CT_LS_8", "A"
            .add "FTTX_TAIL_A_FTTX_CG_048CT_LS_8", "A"
            .add "FTTX_TAIL_A_FTTX_CG_144CT_LS_8", "A"
            .add "FTTX_TAIL_A_FTTX_CG_288CT_LS_8", "A"
            .add "FTTX_TAIL_U_CG_024 CT_LS_8", "U"
            .add "FTTX_TAIL_U_CG_048 CT_LS_8", "U"
            .add "FTTX_TAIL_U_CG_144 CT_LS_8", "U"
            .add "FTTX_TAIL_U_CG_288 CT_LS_8", "U"
            .add "FTTX_TAIL_A_CG_024 CT_LS_8", "A"
            .add "FTTX_TAIL_A_CG_048 CT_LS_8", "A"
            .add "FTTX_TAIL_A_CG_144 CT_LS_8", "A"
            .add "FTTX_TAIL_A_CG_288 CT_LS_8", "A"
            .add "FTTX_TAIL_U_024 CT_8", "U"
            .add "FTTX_TAIL_U_048 CT_8", "U"
            .add "FTTX_TAIL_U_144 CT_8", "U"
            .add "FTTX_TAIL_U_288 CT_8", "U"
            .add "FTTX_TAIL_A_024 CT_8", "A"
            .add "FTTX_TAIL_A_048 CT_8", "A"
            .add "FTTX_TAIL_A_144 CT_8", "A"
            .add "FTTX_TAIL_A_288 CT_8", "A"
        End With
    End If
    Set fiber_tail_8ct = fiber_tail_8ct_p
End Property

' 12CT Tails
Public Property Get fiber_tail_12ct() As Dictionary
    If fiber_tail_12ct_p Is Nothing Then
        Set fiber_tail_12ct_p = New Dictionary
        With fiber_tail_12ct_p
            .add "FTTX_TAIL_U_FTTX_CG_024CT_LS_12", "U"
            .add "FTTX_TAIL_U_FTTX_CG_048CT_LS_12", "U"
            .add "FTTX_TAIL_U_FTTX_CG_144CT_LS_12", "U"
            .add "FTTX_TAIL_U_FTTX_CG_288CT_LS_12", "U"
            .add "FTTX_TAIL_A_FTTX_CG_024CT_LS_12", "A"
            .add "FTTX_TAIL_A_FTTX_CG_048CT_LS_12", "A"
            .add "FTTX_TAIL_A_FTTX_CG_144CT_LS_12", "A"
            .add "FTTX_TAIL_A_FTTX_CG_288CT_LS_12", "A"
            .add "FTTX_TAIL_U_CG_024 CT_LS_12", "U"
            .add "FTTX_TAIL_U_CG_048 CT_LS_12", "U"
            .add "FTTX_TAIL_U_CG_144 CT_LS_12", "U"
            .add "FTTX_TAIL_U_CG_288 CT_LS_12", "U"
            .add "FTTX_TAIL_A_CG_024 CT_LS_12", "A"
            .add "FTTX_TAIL_A_CG_048 CT_LS_12", "A"
            .add "FTTX_TAIL_A_CG_144 CT_LS_12", "A"
            .add "FTTX_TAIL_A_CG_288 CT_LS_12", "A"
            .add "FTTX_TAIL_U_024 CT_12", "U"
            .add "FTTX_TAIL_U_048 CT_12", "U"
            .add "FTTX_TAIL_U_144 CT_12", "U"
            .add "FTTX_TAIL_U_288 CT_12", "U"
            .add "FTTX_TAIL_A_024 CT_12", "A"
            .add "FTTX_TAIL_A_048 CT_12", "A"
            .add "FTTX_TAIL_A_144 CT_12", "A"
            .add "FTTX_TAIL_A_288 CT_12", "A"
        End With
    End If
    Set fiber_tail_12ct = fiber_tail_12ct_p
End Property

' 6CT Fiber
Public Property Get fiber_sheath_6ct() As Dictionary
    If fiber_sheath_6ct_p Is Nothing Then
        Set fiber_sheath_6ct_p = New Dictionary
        With fiber_sheath_6ct_p
            .add "FTTX_DIST_U_FTTX_CG_024CT_LS_6", "" 'NOTE: 'U' and 'A' can be added to these rows to enforce checking the Aerial/UG location in BOMs. This was not included here since Magellan doesn't allow changing models at risers
            .add "FTTX_DIST_U_FTTX_CG_048CT_LS_6", ""
            .add "FTTX_DIST_U_FTTX_CG_144CT_LS_6", ""
            .add "FTTX_DIST_U_FTTX_CG_288CT_LS_6", ""
            .add "FTTX_DIST_A_FTTX_CG_024CT_LS_6", ""
            .add "FTTX_DIST_A_FTTX_CG_048CT_LS_6", ""
            .add "FTTX_DIST_A_FTTX_CG_144CT_LS_6", ""
            .add "FTTX_DIST_A_FTTX_CG_288CT_LS_6", ""
            .add "FTTX_TRUNK_U_FTTX_CG_024CT_LS_6", ""
            .add "FTTX_TRUNK_U_FTTX_CG_048CT_LS_6", ""
            .add "FTTX_TRUNK_U_FTTX_CG_144CT_LS_6", ""
            .add "FTTX_TRUNK_U_FTTX_CG_288CT_LS_6", ""
            .add "FTTX_TRUNK_A_FTTX_CG_024CT_LS_6", ""
            .add "FTTX_TRUNK_A_FTTX_CG_048CT_LS_6", ""
            .add "FTTX_TRUNK_A_FTTX_CG_144CT_LS_6", ""
            .add "FTTX_TRUNK_A_FTTX_CG_288CT_LS_6", ""
            .add "FTTX_DIST_U_CG_024 CT_LS_6", ""
            .add "FTTX_DIST_U_CG_048 CT_LS_6", ""
            .add "FTTX_DIST_U_CG_144 CT_LS_6", ""
            .add "FTTX_DIST_U_CG_288 CT_LS_6", ""
            .add "FTTX_DIST_A_CG_024 CT_LS_6", ""
            .add "FTTX_DIST_A_CG_048 CT_LS_6", ""
            .add "FTTX_DIST_A_CG_144 CT_LS_6", ""
            .add "FTTX_DIST_A_CG_288 CT_LS_6", ""
            .add "FTTX_TRUNK_U_CG_024 CT_LS_6", ""
            .add "FTTX_TRUNK_U_CG_048 CT_LS_6", ""
            .add "FTTX_TRUNK_U_CG_144 CT_LS_6", ""
            .add "FTTX_TRUNK_U_CG_288 CT_LS_6", ""
            .add "FTTX_TRUNK_A_CG_024 CT_LS_6", ""
            .add "FTTX_TRUNK_A_CG_048 CT_LS_6", ""
            .add "FTTX_TRUNK_A_CG_144 CT_LS_6", ""
            .add "FTTX_TRUNK_A_CG_288 CT_LS_6", ""
            .add "FTTX_DIST_U_024 CT_6", ""
            .add "FTTX_DIST_U_048 CT_6", ""
            .add "FTTX_DIST_U_144 CT_6", ""
            .add "FTTX_DIST_U_288 CT_6", ""
            .add "FTTX_DIST_A_024 CT_6", ""
            .add "FTTX_DIST_A_048 CT_6", ""
            .add "FTTX_DIST_A_144 CT_6", ""
            .add "FTTX_DIST_A_288 CT_6", ""
            .add "FTTX_TRUNK_U_024 CT_6", ""
            .add "FTTX_TRUNK_U_048 CT_6", ""
            .add "FTTX_TRUNK_U_144 CT_6", ""
            .add "FTTX_TRUNK_U_288 CT_6", ""
            .add "FTTX_TRUNK_A_024 CT_6", ""
            .add "FTTX_TRUNK_A_048 CT_6", ""
            .add "FTTX_TRUNK_A_144 CT_6", ""
            .add "FTTX_TRUNK_A_288 CT_6", ""
        End With
    End If
    Set fiber_sheath_6ct = fiber_sheath_6ct_p
End Property

' 12CT Fiber
Public Property Get fiber_sheath_12ct() As Dictionary
    If fiber_sheath_12ct_p Is Nothing Then
        Set fiber_sheath_12ct_p = New Dictionary
        With fiber_sheath_12ct_p
            .add "FTTX_DIST_U_FTTX_CG_024CT_LS_12", ""
            .add "FTTX_DIST_U_FTTX_CG_048CT_LS_12", ""
            .add "FTTX_DIST_U_FTTX_CG_144CT_LS_12", ""
            .add "FTTX_DIST_U_FTTX_CG_288CT_LS_12", ""
            .add "FTTX_DIST_A_FTTX_CG_024CT_LS_12", ""
            .add "FTTX_DIST_A_FTTX_CG_048CT_LS_12", ""
            .add "FTTX_DIST_A_FTTX_CG_144CT_LS_12", ""
            .add "FTTX_DIST_A_FTTX_CG_288CT_LS_12", ""
            .add "FTTX_TRUNK_U_FTTX_CG_024CT_LS_12", ""
            .add "FTTX_TRUNK_U_FTTX_CG_048CT_LS_12", ""
            .add "FTTX_TRUNK_U_FTTX_CG_144CT_LS_12", ""
            .add "FTTX_TRUNK_U_FTTX_CG_288CT_LS_12", ""
            .add "FTTX_TRUNK_A_FTTX_CG_024CT_LS_12", ""
            .add "FTTX_TRUNK_A_FTTX_CG_048CT_LS_12", ""
            .add "FTTX_TRUNK_A_FTTX_CG_144CT_LS_12", ""
            .add "FTTX_TRUNK_A_FTTX_CG_288CT_LS_12", ""
            .add "FTTX_DIST_U_CG_024 CT_LS_12", ""
            .add "FTTX_DIST_U_CG_048 CT_LS_12", ""
            .add "FTTX_DIST_U_CG_144 CT_LS_12", ""
            .add "FTTX_DIST_U_CG_288 CT_LS_12", ""
            .add "FTTX_DIST_A_CG_024 CT_LS_12", ""
            .add "FTTX_DIST_A_CG_048 CT_LS_12", ""
            .add "FTTX_DIST_A_CG_144 CT_LS_12", ""
            .add "FTTX_DIST_A_CG_288 CT_LS_12", ""
            .add "FTTX_TRUNK_U_CG_024 CT_LS_12", ""
            .add "FTTX_TRUNK_U_CG_048 CT_LS_12", ""
            .add "FTTX_TRUNK_U_CG_144 CT_LS_12", ""
            .add "FTTX_TRUNK_U_CG_288 CT_LS_12", ""
            .add "FTTX_TRUNK_A_CG_024 CT_LS_12", ""
            .add "FTTX_TRUNK_A_CG_048 CT_LS_12", ""
            .add "FTTX_TRUNK_A_CG_144 CT_LS_12", ""
            .add "FTTX_TRUNK_A_CG_288 CT_LS_12", ""
            .add "FTTX_DIST_U_024 CT_12", ""
            .add "FTTX_DIST_U_048 CT_12", ""
            .add "FTTX_DIST_U_144 CT_12", ""
            .add "FTTX_DIST_U_288 CT_12", ""
            .add "FTTX_DIST_A_024 CT_12", ""
            .add "FTTX_DIST_A_048 CT_12", ""
            .add "FTTX_DIST_A_144 CT_12", ""
            .add "FTTX_DIST_A_288 CT_12", ""
            .add "FTTX_TRUNK_U_024 CT_12", ""
            .add "FTTX_TRUNK_U_048 CT_12", ""
            .add "FTTX_TRUNK_U_144 CT_12", ""
            .add "FTTX_TRUNK_U_288 CT_12", ""
            .add "FTTX_TRUNK_A_024 CT_12", ""
            .add "FTTX_TRUNK_A_048 CT_12", ""
            .add "FTTX_TRUNK_A_144 CT_12", ""
            .add "FTTX_TRUNK_A_288 CT_12", ""
        End With
    End If
    Set fiber_sheath_12ct = fiber_sheath_12ct_p
End Property

' 24CT Fiber
Public Property Get fiber_sheath_24ct() As Dictionary
    If fiber_sheath_24ct_p Is Nothing Then
        Set fiber_sheath_24ct_p = New Dictionary
        With fiber_sheath_24ct_p
            .add "FTTX_DIST_U_FTTX_CG_024CT_LS_24", ""
            .add "FTTX_DIST_U_FTTX_CG_048CT_LS_24", ""
            .add "FTTX_DIST_U_FTTX_CG_144CT_LS_24", ""
            .add "FTTX_DIST_U_FTTX_CG_288CT_LS_24", ""
            .add "FTTX_DIST_A_FTTX_CG_024CT_LS_24", ""
            .add "FTTX_DIST_A_FTTX_CG_048CT_LS_24", ""
            .add "FTTX_DIST_A_FTTX_CG_144CT_LS_24", ""
            .add "FTTX_DIST_A_FTTX_CG_288CT_LS_24", ""
            .add "FTTX_TRUNK_U_FTTX_CG_024CT_LS_24", ""
            .add "FTTX_TRUNK_U_FTTX_CG_048CT_LS_24", ""
            .add "FTTX_TRUNK_U_FTTX_CG_144CT_LS_24", ""
            .add "FTTX_TRUNK_U_FTTX_CG_288CT_LS_24", ""
            .add "FTTX_TRUNK_A_FTTX_CG_024CT_LS_24", ""
            .add "FTTX_TRUNK_A_FTTX_CG_048CT_LS_24", ""
            .add "FTTX_TRUNK_A_FTTX_CG_144CT_LS_24", ""
            .add "FTTX_TRUNK_A_FTTX_CG_288CT_LS_24", ""
            .add "FTTX_DIST_U_CG_024 CT_LS_24", ""
            .add "FTTX_DIST_U_CG_048 CT_LS_24", ""
            .add "FTTX_DIST_U_CG_144 CT_LS_24", ""
            .add "FTTX_DIST_U_CG_288 CT_LS_24", ""
            .add "FTTX_DIST_A_CG_024 CT_LS_24", ""
            .add "FTTX_DIST_A_CG_048 CT_LS_24", ""
            .add "FTTX_DIST_A_CG_144 CT_LS_24", ""
            .add "FTTX_DIST_A_CG_288 CT_LS_24", ""
            .add "FTTX_TRUNK_U_CG_024 CT_LS_24", ""
            .add "FTTX_TRUNK_U_CG_048 CT_LS_24", ""
            .add "FTTX_TRUNK_U_CG_144 CT_LS_24", ""
            .add "FTTX_TRUNK_U_CG_288 CT_LS_24", ""
            .add "FTTX_TRUNK_A_CG_024 CT_LS_24", ""
            .add "FTTX_TRUNK_A_CG_048 CT_LS_24", ""
            .add "FTTX_TRUNK_A_CG_144 CT_LS_24", ""
            .add "FTTX_TRUNK_A_CG_288 CT_LS_24", ""
            .add "FTTX_DIST_U_024 CT_24", ""
            .add "FTTX_DIST_U_048 CT_24", ""
            .add "FTTX_DIST_U_144 CT_24", ""
            .add "FTTX_DIST_U_288 CT_24", ""
            .add "FTTX_DIST_A_024 CT_24", ""
            .add "FTTX_DIST_A_048 CT_24", ""
            .add "FTTX_DIST_A_144 CT_24", ""
            .add "FTTX_DIST_A_288 CT_24", ""
            .add "FTTX_TRUNK_U_024 CT_24", ""
            .add "FTTX_TRUNK_U_048 CT_24", ""
            .add "FTTX_TRUNK_U_144 CT_24", ""
            .add "FTTX_TRUNK_U_288 CT_24", ""
            .add "FTTX_TRUNK_A_024 CT_24", ""
            .add "FTTX_TRUNK_A_048 CT_24", ""
            .add "FTTX_TRUNK_A_144 CT_24", ""
            .add "FTTX_TRUNK_A_288 CT_24", ""
        End With
    End If
    Set fiber_sheath_24ct = fiber_sheath_24ct_p
End Property

' 48CT Fiber
Public Property Get fiber_sheath_48ct() As Dictionary
    If fiber_sheath_48ct_p Is Nothing Then
        Set fiber_sheath_48ct_p = New Dictionary
        With fiber_sheath_48ct_p
            .add "FTTX_DIST_U_FTTX_CG_024CT_LS_48", ""
            .add "FTTX_DIST_U_FTTX_CG_048CT_LS_48", ""
            .add "FTTX_DIST_U_FTTX_CG_144CT_LS_48", ""
            .add "FTTX_DIST_U_FTTX_CG_288CT_LS_48", ""
            .add "FTTX_DIST_A_FTTX_CG_024CT_LS_48", ""
            .add "FTTX_DIST_A_FTTX_CG_048CT_LS_48", ""
            .add "FTTX_DIST_A_FTTX_CG_144CT_LS_48", ""
            .add "FTTX_DIST_A_FTTX_CG_288CT_LS_48", ""
            .add "FTTX_TRUNK_U_FTTX_CG_024CT_LS_48", ""
            .add "FTTX_TRUNK_U_FTTX_CG_048CT_LS_48", ""
            .add "FTTX_TRUNK_U_FTTX_CG_144CT_LS_48", ""
            .add "FTTX_TRUNK_U_FTTX_CG_288CT_LS_48", ""
            .add "FTTX_TRUNK_A_FTTX_CG_024CT_LS_48", ""
            .add "FTTX_TRUNK_A_FTTX_CG_048CT_LS_48", ""
            .add "FTTX_TRUNK_A_FTTX_CG_144CT_LS_48", ""
            .add "FTTX_TRUNK_A_FTTX_CG_288CT_LS_48", ""
            .add "FTTX_DIST_U_CG_024 CT_LS_48", ""
            .add "FTTX_DIST_U_CG_048 CT_LS_48", ""
            .add "FTTX_DIST_U_CG_144 CT_LS_48", ""
            .add "FTTX_DIST_U_CG_288 CT_LS_48", ""
            .add "FTTX_DIST_A_CG_024 CT_LS_48", ""
            .add "FTTX_DIST_A_CG_048 CT_LS_48", ""
            .add "FTTX_DIST_A_CG_144 CT_LS_48", ""
            .add "FTTX_DIST_A_CG_288 CT_LS_48", ""
            .add "FTTX_TRUNK_U_CG_024 CT_LS_48", ""
            .add "FTTX_TRUNK_U_CG_048 CT_LS_48", ""
            .add "FTTX_TRUNK_U_CG_144 CT_LS_48", ""
            .add "FTTX_TRUNK_U_CG_288 CT_LS_48", ""
            .add "FTTX_TRUNK_A_CG_024 CT_LS_48", ""
            .add "FTTX_TRUNK_A_CG_048 CT_LS_48", ""
            .add "FTTX_TRUNK_A_CG_144 CT_LS_48", ""
            .add "FTTX_TRUNK_A_CG_288 CT_LS_48", ""
            .add "FTTX_DIST_U_024 CT_48", ""
            .add "FTTX_DIST_U_048 CT_48", ""
            .add "FTTX_DIST_U_144 CT_48", ""
            .add "FTTX_DIST_U_288 CT_48", ""
            .add "FTTX_DIST_A_024 CT_48", ""
            .add "FTTX_DIST_A_048 CT_48", ""
            .add "FTTX_DIST_A_144 CT_48", ""
            .add "FTTX_DIST_A_288 CT_48", ""
            .add "FTTX_TRUNK_U_024 CT_48", ""
            .add "FTTX_TRUNK_U_048 CT_48", ""
            .add "FTTX_TRUNK_U_144 CT_48", ""
            .add "FTTX_TRUNK_U_288 CT_48", ""
            .add "FTTX_TRUNK_A_024 CT_48", ""
            .add "FTTX_TRUNK_A_048 CT_48", ""
            .add "FTTX_TRUNK_A_144 CT_48", ""
            .add "FTTX_TRUNK_A_288 CT_48", ""
        End With
    End If
    Set fiber_sheath_48ct = fiber_sheath_48ct_p
End Property

' 72CT Fiber
Public Property Get fiber_sheath_72ct() As Dictionary
    If fiber_sheath_72ct_p Is Nothing Then
        Set fiber_sheath_72ct_p = New Dictionary
        With fiber_sheath_72ct_p
            .add "FTTX_DIST_U_FTTX_CG_024CT_LS_72", ""
            .add "FTTX_DIST_U_FTTX_CG_048CT_LS_72", ""
            .add "FTTX_DIST_U_FTTX_CG_144CT_LS_72", ""
            .add "FTTX_DIST_U_FTTX_CG_288CT_LS_72", ""
            .add "FTTX_DIST_A_FTTX_CG_024CT_LS_72", ""
            .add "FTTX_DIST_A_FTTX_CG_048CT_LS_72", ""
            .add "FTTX_DIST_A_FTTX_CG_144CT_LS_72", ""
            .add "FTTX_DIST_A_FTTX_CG_288CT_LS_72", ""
            .add "FTTX_TRUNK_U_FTTX_CG_024CT_LS_72", ""
            .add "FTTX_TRUNK_U_FTTX_CG_048CT_LS_72", ""
            .add "FTTX_TRUNK_U_FTTX_CG_144CT_LS_72", ""
            .add "FTTX_TRUNK_U_FTTX_CG_288CT_LS_72", ""
            .add "FTTX_TRUNK_A_FTTX_CG_024CT_LS_72", ""
            .add "FTTX_TRUNK_A_FTTX_CG_048CT_LS_72", ""
            .add "FTTX_TRUNK_A_FTTX_CG_144CT_LS_72", ""
            .add "FTTX_TRUNK_A_FTTX_CG_288CT_LS_72", ""
            .add "FTTX_DIST_U_CG_024 CT_LS_72", ""
            .add "FTTX_DIST_U_CG_048 CT_LS_72", ""
            .add "FTTX_DIST_U_CG_144 CT_LS_72", ""
            .add "FTTX_DIST_U_CG_288 CT_LS_72", ""
            .add "FTTX_DIST_A_CG_024 CT_LS_72", ""
            .add "FTTX_DIST_A_CG_048 CT_LS_72", ""
            .add "FTTX_DIST_A_CG_144 CT_LS_72", ""
            .add "FTTX_DIST_A_CG_288 CT_LS_72", ""
            .add "FTTX_TRUNK_U_CG_024 CT_LS_72", ""
            .add "FTTX_TRUNK_U_CG_048 CT_LS_72", ""
            .add "FTTX_TRUNK_U_CG_144 CT_LS_72", ""
            .add "FTTX_TRUNK_U_CG_288 CT_LS_72", ""
            .add "FTTX_TRUNK_A_CG_024 CT_LS_72", ""
            .add "FTTX_TRUNK_A_CG_048 CT_LS_72", ""
            .add "FTTX_TRUNK_A_CG_144 CT_LS_72", ""
            .add "FTTX_TRUNK_A_CG_288 CT_LS_72", ""
            .add "FTTX_DIST_U_024 CT_72", ""
            .add "FTTX_DIST_U_048 CT_72", ""
            .add "FTTX_DIST_U_144 CT_72", ""
            .add "FTTX_DIST_U_288 CT_72", ""
            .add "FTTX_DIST_A_024 CT_72", ""
            .add "FTTX_DIST_A_048 CT_72", ""
            .add "FTTX_DIST_A_144 CT_72", ""
            .add "FTTX_DIST_A_288 CT_72", ""
            .add "FTTX_TRUNK_U_024 CT_72", ""
            .add "FTTX_TRUNK_U_048 CT_72", ""
            .add "FTTX_TRUNK_U_144 CT_72", ""
            .add "FTTX_TRUNK_U_288 CT_72", ""
            .add "FTTX_TRUNK_A_024 CT_72", ""
            .add "FTTX_TRUNK_A_048 CT_72", ""
            .add "FTTX_TRUNK_A_144 CT_72", ""
            .add "FTTX_TRUNK_A_288 CT_72", ""
        End With
    End If
    Set fiber_sheath_72ct = fiber_sheath_72ct_p
End Property

' 96CT Fiber
Public Property Get fiber_sheath_96ct() As Dictionary
    If fiber_sheath_96ct_p Is Nothing Then
        Set fiber_sheath_96ct_p = New Dictionary
        With fiber_sheath_96ct_p
            .add "FTTX_DIST_U_FTTX_CG_024CT_LS_96", ""
            .add "FTTX_DIST_U_FTTX_CG_048CT_LS_96", ""
            .add "FTTX_DIST_U_FTTX_CG_144CT_LS_96", ""
            .add "FTTX_DIST_U_FTTX_CG_288CT_LS_96", ""
            .add "FTTX_DIST_A_FTTX_CG_024CT_LS_96", ""
            .add "FTTX_DIST_A_FTTX_CG_048CT_LS_96", ""
            .add "FTTX_DIST_A_FTTX_CG_144CT_LS_96", ""
            .add "FTTX_DIST_A_FTTX_CG_288CT_LS_96", ""
            .add "FTTX_TRUNK_U_FTTX_CG_024CT_LS_96", ""
            .add "FTTX_TRUNK_U_FTTX_CG_048CT_LS_96", ""
            .add "FTTX_TRUNK_U_FTTX_CG_144CT_LS_96", ""
            .add "FTTX_TRUNK_U_FTTX_CG_288CT_LS_96", ""
            .add "FTTX_TRUNK_A_FTTX_CG_024CT_LS_96", ""
            .add "FTTX_TRUNK_A_FTTX_CG_048CT_LS_96", ""
            .add "FTTX_TRUNK_A_FTTX_CG_144CT_LS_96", ""
            .add "FTTX_TRUNK_A_FTTX_CG_288CT_LS_96", ""
            .add "FTTX_DIST_U_CG_024 CT_LS_96", ""
            .add "FTTX_DIST_U_CG_048 CT_LS_96", ""
            .add "FTTX_DIST_U_CG_144 CT_LS_96", ""
            .add "FTTX_DIST_U_CG_288 CT_LS_96", ""
            .add "FTTX_DIST_A_CG_024 CT_LS_96", ""
            .add "FTTX_DIST_A_CG_048 CT_LS_96", ""
            .add "FTTX_DIST_A_CG_144 CT_LS_96", ""
            .add "FTTX_DIST_A_CG_288 CT_LS_96", ""
            .add "FTTX_TRUNK_U_CG_024 CT_LS_96", ""
            .add "FTTX_TRUNK_U_CG_048 CT_LS_96", ""
            .add "FTTX_TRUNK_U_CG_144 CT_LS_96", ""
            .add "FTTX_TRUNK_U_CG_288 CT_LS_96", ""
            .add "FTTX_TRUNK_A_CG_024 CT_LS_96", ""
            .add "FTTX_TRUNK_A_CG_048 CT_LS_96", ""
            .add "FTTX_TRUNK_A_CG_144 CT_LS_96", ""
            .add "FTTX_TRUNK_A_CG_288 CT_LS_96", ""
            .add "FTTX_DIST_U_024 CT_96", ""
            .add "FTTX_DIST_U_048 CT_96", ""
            .add "FTTX_DIST_U_144 CT_96", ""
            .add "FTTX_DIST_U_288 CT_96", ""
            .add "FTTX_DIST_A_024 CT_96", ""
            .add "FTTX_DIST_A_048 CT_96", ""
            .add "FTTX_DIST_A_144 CT_96", ""
            .add "FTTX_DIST_A_288 CT_96", ""
            .add "FTTX_TRUNK_U_024 CT_96", ""
            .add "FTTX_TRUNK_U_048 CT_96", ""
            .add "FTTX_TRUNK_U_144 CT_96", ""
            .add "FTTX_TRUNK_U_288 CT_96", ""
            .add "FTTX_TRUNK_A_024 CT_96", ""
            .add "FTTX_TRUNK_A_048 CT_96", ""
            .add "FTTX_TRUNK_A_144 CT_96", ""
            .add "FTTX_TRUNK_A_288 CT_96", ""
        End With
    End If
    Set fiber_sheath_96ct = fiber_sheath_96ct_p
End Property

' 144CT Fiber
Public Property Get fiber_sheath_144ct() As Dictionary
    If fiber_sheath_144ct_p Is Nothing Then
        Set fiber_sheath_144ct_p = New Dictionary
        With fiber_sheath_144ct_p
            .add "FTTX_DIST_U_FTTX_CG_024CT_LS_144", ""
            .add "FTTX_DIST_U_FTTX_CG_048CT_LS_144", ""
            .add "FTTX_DIST_U_FTTX_CG_144CT_LS_144", ""
            .add "FTTX_DIST_U_FTTX_CG_288CT_LS_144", ""
            .add "FTTX_DIST_A_FTTX_CG_024CT_LS_144", ""
            .add "FTTX_DIST_A_FTTX_CG_048CT_LS_144", ""
            .add "FTTX_DIST_A_FTTX_CG_144CT_LS_144", ""
            .add "FTTX_DIST_A_FTTX_CG_288CT_LS_144", ""
            .add "FTTX_TRUNK_U_FTTX_CG_024CT_LS_144", ""
            .add "FTTX_TRUNK_U_FTTX_CG_048CT_LS_144", ""
            .add "FTTX_TRUNK_U_FTTX_CG_144CT_LS_144", ""
            .add "FTTX_TRUNK_U_FTTX_CG_288CT_LS_144", ""
            .add "FTTX_TRUNK_A_FTTX_CG_024CT_LS_144", ""
            .add "FTTX_TRUNK_A_FTTX_CG_048CT_LS_144", ""
            .add "FTTX_TRUNK_A_FTTX_CG_144CT_LS_144", ""
            .add "FTTX_TRUNK_A_FTTX_CG_288CT_LS_144", ""
            .add "FTTX_DIST_U_CG_024 CT_LS_144", ""
            .add "FTTX_DIST_U_CG_048 CT_LS_144", ""
            .add "FTTX_DIST_U_CG_144 CT_LS_144", ""
            .add "FTTX_DIST_U_CG_288 CT_LS_144", ""
            .add "FTTX_DIST_A_CG_024 CT_LS_144", ""
            .add "FTTX_DIST_A_CG_048 CT_LS_144", ""
            .add "FTTX_DIST_A_CG_144 CT_LS_144", ""
            .add "FTTX_DIST_A_CG_288 CT_LS_144", ""
            .add "FTTX_TRUNK_U_CG_024 CT_LS_144", ""
            .add "FTTX_TRUNK_U_CG_048 CT_LS_144", ""
            .add "FTTX_TRUNK_U_CG_144 CT_LS_144", ""
            .add "FTTX_TRUNK_U_CG_288 CT_LS_144", ""
            .add "FTTX_TRUNK_A_CG_024 CT_LS_144", ""
            .add "FTTX_TRUNK_A_CG_048 CT_LS_144", ""
            .add "FTTX_TRUNK_A_CG_144 CT_LS_144", ""
            .add "FTTX_TRUNK_A_CG_288 CT_LS_144", ""
            .add "FTTX_DIST_U_024 CT_144", ""
            .add "FTTX_DIST_U_048 CT_144", ""
            .add "FTTX_DIST_U_144 CT_144", ""
            .add "FTTX_DIST_U_288 CT_144", ""
            .add "FTTX_DIST_A_024 CT_144", ""
            .add "FTTX_DIST_A_048 CT_144", ""
            .add "FTTX_DIST_A_144 CT_144", ""
            .add "FTTX_DIST_A_288 CT_144", ""
            .add "FTTX_TRUNK_U_024 CT_144", ""
            .add "FTTX_TRUNK_U_048 CT_144", ""
            .add "FTTX_TRUNK_U_144 CT_144", ""
            .add "FTTX_TRUNK_U_288 CT_144", ""
            .add "FTTX_TRUNK_A_024 CT_144", ""
            .add "FTTX_TRUNK_A_048 CT_144", ""
            .add "FTTX_TRUNK_A_144 CT_144", ""
            .add "FTTX_TRUNK_A_288 CT_144", ""
        End With
    End If
    Set fiber_sheath_144ct = fiber_sheath_144ct_p
End Property

' 288CT Fiber
Public Property Get fiber_sheath_288ct() As Dictionary
    If fiber_sheath_288ct_p Is Nothing Then
        Set fiber_sheath_288ct_p = New Dictionary
        With fiber_sheath_288ct_p
            .add "FTTX_DIST_U_FTTX_CG_024CT_LS_288", ""
            .add "FTTX_DIST_U_FTTX_CG_048CT_LS_288", ""
            .add "FTTX_DIST_U_FTTX_CG_144CT_LS_288", ""
            .add "FTTX_DIST_U_FTTX_CG_288CT_LS_288", ""
            .add "FTTX_DIST_A_FTTX_CG_024CT_LS_288", ""
            .add "FTTX_DIST_A_FTTX_CG_048CT_LS_288", ""
            .add "FTTX_DIST_A_FTTX_CG_144CT_LS_288", ""
            .add "FTTX_DIST_A_FTTX_CG_288CT_LS_288", ""
            .add "FTTX_TRUNK_U_FTTX_CG_024CT_LS_288", ""
            .add "FTTX_TRUNK_U_FTTX_CG_048CT_LS_288", ""
            .add "FTTX_TRUNK_U_FTTX_CG_144CT_LS_288", ""
            .add "FTTX_TRUNK_U_FTTX_CG_288CT_LS_288", ""
            .add "FTTX_TRUNK_A_FTTX_CG_024CT_LS_288", ""
            .add "FTTX_TRUNK_A_FTTX_CG_048CT_LS_288", ""
            .add "FTTX_TRUNK_A_FTTX_CG_144CT_LS_288", ""
            .add "FTTX_TRUNK_A_FTTX_CG_288CT_LS_288", ""
            .add "FTTX_DIST_U_CG_024 CT_LS_288", ""
            .add "FTTX_DIST_U_CG_048 CT_LS_288", ""
            .add "FTTX_DIST_U_CG_144 CT_LS_288", ""
            .add "FTTX_DIST_U_CG_288 CT_LS_288", ""
            .add "FTTX_DIST_A_CG_024 CT_LS_288", ""
            .add "FTTX_DIST_A_CG_048 CT_LS_288", ""
            .add "FTTX_DIST_A_CG_144 CT_LS_288", ""
            .add "FTTX_DIST_A_CG_288 CT_LS_288", ""
            .add "FTTX_TRUNK_U_CG_024 CT_LS_288", ""
            .add "FTTX_TRUNK_U_CG_048 CT_LS_288", ""
            .add "FTTX_TRUNK_U_CG_144 CT_LS_288", ""
            .add "FTTX_TRUNK_U_CG_288 CT_LS_288", ""
            .add "FTTX_TRUNK_A_CG_024 CT_LS_288", ""
            .add "FTTX_TRUNK_A_CG_048 CT_LS_288", ""
            .add "FTTX_TRUNK_A_CG_144 CT_LS_288", ""
            .add "FTTX_TRUNK_A_CG_288 CT_LS_288", ""
            .add "FTTX_DIST_U_024 CT_288", ""
            .add "FTTX_DIST_U_048 CT_288", ""
            .add "FTTX_DIST_U_144 CT_288", ""
            .add "FTTX_DIST_U_288 CT_288", ""
            .add "FTTX_DIST_A_024 CT_288", ""
            .add "FTTX_DIST_A_048 CT_288", ""
            .add "FTTX_DIST_A_144 CT_288", ""
            .add "FTTX_DIST_A_288 CT_288", ""
            .add "FTTX_TRUNK_U_024 CT_288", ""
            .add "FTTX_TRUNK_U_048 CT_288", ""
            .add "FTTX_TRUNK_U_144 CT_288", ""
            .add "FTTX_TRUNK_U_288 CT_288", ""
            .add "FTTX_TRUNK_A_024 CT_288", ""
            .add "FTTX_TRUNK_A_048 CT_288", ""
            .add "FTTX_TRUNK_A_144 CT_288", ""
            .add "FTTX_TRUNK_A_288 CT_288", ""
        End With
    End If
    Set fiber_sheath_288ct = fiber_sheath_288ct_p
End Property
