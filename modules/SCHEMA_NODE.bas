Attribute VB_Name = "SCHEMA_NODE"
Public Const NODE_HEADER_ROW = 1

Enum BOM_NODE_COLS
        NODE_POLYGON
        NODE_SPECFILE
        NODE_MFG
        NODE_MAKE
        NODE_MODEL
        NODE_COUNT
        NODE_CONFIG_TYPE
        NODE_CLASSIFICATION
        NODE_STATE_ASBUILT
        NODE_DESIGN
        NODE_NOT_BUILT
        NODE_UPGRADE
End Enum

Private conversion_types_p As Dictionary
Private bom_node_cols_dict_p As Dictionary

Private Property Get bom_node_cols_dict() As Dictionary
        If bom_node_cols_dict_p Is Nothing Then
                Set bom_node_cols_dict_p = New Dictionary
                With bom_node_cols_dict_p
                        .add "POLYGON", NODE_POLYGON
                        .add "SPECFILE", NODE_SPECFILE
                        .add "MFG", NODE_MFG
                        .add "MAKE", NODE_MAKE
                        .add "MODEL", NODE_MODEL
                        .add "COUNT", NODE_COUNT
                        .add "CONFIG_TYPE", NODE_CONFIG_TYPE
                        .add "CLASSIFICATION", NODE_CLASSIFICATION
                        .add "ASBUILT", NODE_STATE_ASBUILT
                        .add "DESIGN", NODE_DESIGN
                        .add "NOT BUILT", NODE_NOT_BUILT
                        .add "UPGRADE", NODE_UPGRADE
                End With
        End If
        Set bom_node_cols_dict = bom_node_cols_dict_p
End Property

Public Function bom_node_cols_from_string(ByVal bom_node_cols_text As String) As BOM_NODE_COLS
        bom_node_cols_from_string = enum_from_string(bom_node_cols_text, bom_node_cols_dict)
End Function

Public Function bom_node_cols_to_string(ByVal bom_node_cols_enum As BOM_NODE_COLS) As String
        bom_node_cols_to_string = enum_to_string(bom_node_cols_enum, bom_node_cols_dict)
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
        Set get_column_map = SCHEMA.get_column_map(sheet, NODE_HEADER_ROW, bom_node_cols_dict)
End Function

Public Function apply_conversion(ByVal col_name As BOM_NODE_COLS, ByVal in_value As String) As Variant
        Dim result As Variant
        result = Array(SCHEMA.apply_conversion(col_name, in_value, conversion_types))
        If IsObject(result(0)) Then
                Set apply_conversion = result(0)
        Else
                Let apply_conversion = result(0)
        End If
End Function
