Attribute VB_Name = "SCHEMA_INT"
Public Const INT_HEADER_ROW = 1

Enum BOM_INTERNAL_COLS
        INT_POLYGON
        INT_SPECFILE
        INT_EQUIP_TYPE
        INT_EQUIP_MFG
        INT_EQUIP_MAKE
        INT_EQUIP_MODEL
        INT_EQUIP_ID
        INT_EQUIP_NAME
        INT_EQUIP_CLASSIFICATION
        INT_MFG
        INT_MAKE
        INT_MODEL
        INT_NUM_PORTS
        INT_COUNT
        INT_CLASSIFICATION
        INT_STATE_ASBUILT
        INT_STATE_DESIGN
        INT_STATE_NOT_BUILT
        INT_STATE_UPGRADE
End Enum

Private conversion_types_p As Dictionary
Private bom_internal_cols_dict_p As Dictionary

Private Property Get bom_internal_cols_dict() As Dictionary
        If bom_internal_cols_dict_p Is Nothing Then
                Set bom_internal_cols_dict_p = New Dictionary
                With bom_internal_cols_dict_p
                        .add "POLYGON", INT_POLYGON
                        .add "SPECFILE", INT_SPECFILE
                        .add "EQUIP_TYPE", INT_EQUIP_TYPE
                        .add "EQUIP_MFG", INT_EQUIP_MFG
                        .add "EQUIP_MAKE", INT_EQUIP_MAKE
                        .add "EQUIP_MODEL", INT_EQUIP_MODEL
                        .add "EQUIP_ID", INT_EQUIP_ID
                        .add "EQUIP_NAME", INT_EQUIP_NAME
                        .add "EQUIP_CLASSIFICATION", INT_EQUIP_CLASSIFICATION
                        .add "MFG", INT_MFG
                        .add "MAKE", INT_MAKE
                        .add "MODEL", INT_MODEL
                        .add "NUM_PORTS", INT_NUM_PORTS
                        .add "COUNT", INT_COUNT
                        .add "CLASSIFICATION", INT_CLASSIFICATION
                        .add "ASBUILT", INT_STATE_ASBUILT
                        .add "DESIGN", INT_STATE_DESIGN
                        .add "NOT BUILT", INT_STATE_NOT_BUILT
                        .add "UPGRADE", INT_STATE_UPGRADE
                End With
        End If
        Set bom_internal_cols_dict = bom_internal_cols_dict_p
End Property

Public Function bom_internal_cols_from_string(ByVal bom_internal_cols_text As String) As BOM_INTERNAL_COLS
        bom_internal_cols_from_string = enum_from_string(bom_internal_cols_text, bom_internal_cols_dict)
End Function

Public Function bom_internal_cols_to_string(ByVal bom_internal_cols_enum As BOM_INTERNAL_COLS) As String
        bom_internal_cols_to_string = enum_to_string(bom_internal_cols_enum, bom_internal_cols_dict)
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
        Set get_column_map = SCHEMA.get_column_map(sheet, INT_HEADER_ROW, bom_internal_cols_dict)
End Function

Public Function apply_conversion(ByVal col_name As BOM_INTERNAL_COLS, ByVal in_value As String) As Variant
        Dim result As Variant
        result = Array(SCHEMA.apply_conversion(col_name, in_value, conversion_types))
        If IsObject(result(0)) Then
                Set apply_conversion = result(0)
        Else
                Let apply_conversion = result(0)
        End If
End Function
