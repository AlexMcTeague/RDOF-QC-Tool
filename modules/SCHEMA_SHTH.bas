Attribute VB_Name = "SCHEMA_SHTH"
Public Const SHTH_HEADER_ROW = 1

Enum BOM_SHEATH_COLS
        SHTH_POLYGON
        SHTH_LOCATION
        SHTH_MAKE
        SHTH_MODEL
        SHTH_FTG
        SHTH_SLACK_FTG
        SHTH_SLACK_COUNT
        SHTH_TOTAL_FTG
        SHTH_MILES
        SHTH_SLACK_MILES
        SHTH_TOTAL_MILES
        SHTH_CLASSIFICATION
        SHTH_STATE_ASBUILT
        SHTH_STATE_DESIGN
        SHTH_STATE_NOT_BUILT
        SHTH_STATE_UPGRADE
End Enum

Private conversion_types_p As Dictionary
Private bom_sheath_cols_dict_p As Dictionary

Private Property Get bom_sheath_cols_dict() As Dictionary
        If bom_sheath_cols_dict_p Is Nothing Then
                Set bom_sheath_cols_dict_p = New Dictionary
                With bom_sheath_cols_dict_p
                        .add "POLYGON", SHTH_POLYGON
                        .add "LOCATION", SHTH_LOCATION
                        .add "MAKE", SHTH_MAKE
                        .add "MODEL", SHTH_MODEL
                        .add "FTG", SHTH_FTG
                        .add "SLACK_FTG", SHTH_SLACK_FTG
                        .add "SLACK_COUNT", SHTH_SLACK_COUNT
                        .add "TOTAL_FTG", SHTH_TOTAL_FTG
                        .add "MILES", SHTH_MILES
                        .add "SLACK_MILES", SHTH_SLACK_MILES
                        .add "TOTAL_MILES", SHTH_TOTAL_MILES
                        .add "CLASSIFICATION", SHTH_CLASSIFICATION
                        .add "ASBUILT", SHTH_STATE_ASBUILT
                        .add "DESIGN", SHTH_STATE_DESIGN
                        .add "NOT BUILT", SHTH_STATE_NOT_BUILT
                        .add "UPGRADE", SHTH_STATE_UPGRADE
                End With
        End If
        Set bom_sheath_cols_dict = bom_sheath_cols_dict_p
End Property

Public Function bom_sheath_cols_from_string(ByVal bom_sheath_cols_text As String) As BOM_SHEATH_COLS
        bom_sheath_cols_from_string = enum_from_string(bom_sheath_cols_text, bom_sheath_cols_dict)
End Function

Public Function bom_sheath_cols_to_string(ByVal bom_sheath_cols_enum As BOM_SHEATH_COLS) As String
        bom_sheath_cols_to_string = enum_to_string(bom_sheath_cols_enum, bom_sheath_cols_dict)
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
        Set get_column_map = SCHEMA.get_column_map(sheet, SHTH_HEADER_ROW, bom_sheath_cols_dict)
End Function

Public Function apply_conversion(ByVal col_name As BOM_SHEATH_COLS, ByVal in_value As String) As Variant
        Dim result As Variant
        result = Array(SCHEMA.apply_conversion(col_name, in_value, conversion_types))
        If IsObject(result(0)) Then
                Set apply_conversion = result(0)
        Else
                Let apply_conversion = result(0)
        End If
End Function
