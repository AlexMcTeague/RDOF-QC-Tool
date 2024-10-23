Attribute VB_Name = "SCHEMA_SPL"
Public Const SPL_HEADER_ROW = 1

Enum BOM_SPLIT_COLS
        SPL_POLYGON
        SPL_MFG
        SPL_MAKE
        SPL_MODEL
        SPL_COUNT
        SPL_CLASSIFICATION
        SPL_STATE_ASBUILT
        SPL_STATE_DESIGN
        SPL_STATE_NOT_BUILT
        SPL_STATE_UPGRADE
End Enum

Private conversion_types_p As Dictionary
Private bom_split_cols_dict_p As Dictionary

Private Property Get bom_split_cols_dict() As Dictionary
        If bom_split_cols_dict_p Is Nothing Then
                Set bom_split_cols_dict_p = New Dictionary
                With bom_split_cols_dict_p
                        .add "POLYGON", SPL_POLYGON
                        .add "MFG", SPL_MFG
                        .add "MAKE", SPL_MAKE
                        .add "MODEL", SPL_MODEL
                        .add "COUNT", SPL_COUNT
                        .add "CLASSIFICATION", SPL_CLASSIFICATION
                        .add "ASBUILT", SPL_STATE_ASBUILT
                        .add "DESIGN", SPL_STATE_DESIGN
                        .add "NOT BUILT", SPL_STATE_NOT_BUILT
                        .add "UPGRADE", SPL_STATE_UPGRADE
                End With
        End If
        Set bom_split_cols_dict = bom_split_cols_dict_p
End Property

Public Function bom_split_cols_from_string(ByVal bom_split_cols_text As String) As BOM_SPLIT_COLS
        bom_split_cols_from_string = enum_from_string(bom_split_cols_text, bom_split_cols_dict)
End Function

Public Function bom_split_cols_to_string(ByVal bom_split_cols_enum As BOM_SPLIT_COLS) As String
        bom_split_cols_to_string = enum_to_string(bom_split_cols_enum, bom_split_cols_dict)
End Function

Private Property Get conversion_types() As Dictionary
        If conversion_types_p Is Nothing Then
                Set conversion_types_p = New Dictionary
                With conversion_types_p
                End With
        End If
        Set conversion_types = conversion_types_p
End Property

Public Function get_column_map(ByRef sheet As Worksheet) As Dictionary
        Set get_column_map = SCHEMA.get_column_map(sheet, SPL_HEADER_ROW, bom_split_cols_dict)
End Function

Public Function apply_conversion(ByVal col_name As BOM_SPLIT_COLS, ByVal in_value As String) As Variant
        Dim result As Variant
        result = Array(SCHEMA.apply_conversion(col_name, in_value, conversion_types))
        If IsObject(result(0)) Then
                Set apply_conversion = result(0)
        Else
                Let apply_conversion = result(0)
        End If
End Function
