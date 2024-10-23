Attribute VB_Name = "SCHEMA"
Public Function enum_from_string(ByVal in_text As String, ByRef ref_dict As Dictionary) As Variant
    result = -1
    
    If ref_dict.Exists(in_text) Then
        result = ref_dict(in_text)
    End If
    
    enum_from_string = result
End Function

Public Function enum_to_string(ByVal in_enum As Variant, ByRef ref_dict As Dictionary) As String
    result = "ERROR"

    For Each dict_text In ref_dict.Keys
        dict_enum = ref_dict(dict_text)
        
        If dict_enum = in_enum Then
            result = dict_text
            Exit For
        End If
    Next
    
    enum_to_string = result
End Function

Public Function get_column_map(ByRef sheet As Worksheet, ByVal header_row As Long, ByRef ref_dict As Dictionary) As Dictionary
    Dim result As Dictionary
    Set result = New Dictionary
    last_col = sheet.UsedRange.Columns.Count
    
    col_index = 0
    Do While col_index < last_col
        col_text = sheet.Cells(header_row, col_index + 1).Value
        If ref_dict.Exists(col_text) And col_text <> "" Then
            If TypeName(ref_dict(col_text)) = "Variant()" Then
                disam = disambiguate_default(ref_dict(col_text), result, sheet)
                If Not IsEmpty(disam) Then
                    result.add disam, Replace(Split(sheet.Cells(header_row, col_index + 1).Address, "$")(1), "$", "")
                End If
            Else
                result.add ref_dict(col_text), Replace(Split(sheet.Cells(header_row, col_index + 1).Address, "$")(1), "$", "")
            End If
        End If
        
        col_index = col_index + 1
    Loop
    
    Set get_column_map = result
End Function

Public Function apply_conversion(ByVal col_name As Variant, ByVal in_value As String, ByRef col_types As Dictionary) As Variant
    Dim result As Variant
    
    If col_types.Exists(col_name) Then
        result = Array(Application.Run(col_types(col_name), in_value))
    Else
        result = Array(in_value)
    End If
    
    If IsObject(result(0)) Then
        Set apply_conversion = result(0)
    Else
        Let apply_conversion = result(0)
    End If
End Function

Function disambiguate_default(ByRef cols As Variant, ByRef current_map As Dictionary, ByRef current_sheet As Worksheet) As Variant
    Dim result As Variant
    result = Empty
    
    For i = LBound(cols) To UBound(cols)
        If Not current_map.Exists(cols(i)) Then
            result = cols(i)
            Exit For
        End If
    Next
    
    disambiguate_default = result
End Function
