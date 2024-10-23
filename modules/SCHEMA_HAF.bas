Attribute VB_Name = "SCHEMA_HAF"
Public Const HAF_HEADER_ROW = 1

Enum HAF_COLS
        HAF_HOUSE_NUMBER
        HAF_HOUSE_FRACTION
        HAF_PRE_DIRECTION
        HAF_STREET_NAME
        HAF_STREET_TYPE
        HAF_POST_DIRECTION
        HAF_SUB_DIVISION
        HAF_BUILDING
        HAF_UNIT_TYPE
        HAF_UNIT_NO
        HAF_LOT_ID
        HAF_CITY_NAME
        HAF_STATE_CODE
        HAF_ZIP_CODE
        HAF_HOOKUP_TYPE
        HAF_DWELLING_TYPE
        HAF_STATUS
        HAF_SERVICEABILITY_CODE
        HAF_INSTALLATION_TYPE
        HAF_NYS
        HAF_NYSBO
        HAF_NODE
        HAF_COMMENT
        HAF_HOUSE_KEY
        HAF_AMP
        HAF_POWER_SUPPLY
        HAF_LAT
        HAF_LONG
        HAF_CBG
        HAF_AWARD_TYPE
        HAF_DROP_LENGTH
        HAF_SIK_READY
        HAF_POLE_PORT_NUMS
End Enum

Private conversion_types_p As Dictionary
Private haf_cols_dict_p As Dictionary

Private Property Get haf_cols_dict() As Dictionary
        If haf_cols_dict_p Is Nothing Then
                Set haf_cols_dict_p = New Dictionary
                With haf_cols_dict_p
                        .add "HOUSE NUMBER", HAF_HOUSE_NUMBER
                        .add "HOUSE FRACTION", HAF_HOUSE_FRACTION
                        .add "PRE DIRECTION", HAF_PRE_DIRECTION
                        .add "STREET NAME", HAF_STREET_NAME
                        .add "STREET TYPE", HAF_STREET_TYPE
                        .add "POST DIRECTION", HAF_POST_DIRECTION
                        .add "SUB DIVISION", HAF_SUB_DIVISION
                        .add "BUILDING", HAF_BUILDING
                        .add "UNIT TYPE", HAF_UNIT_TYPE
                        .add "UNIT NO", HAF_UNIT_NO
                        .add "LOT ID", HAF_LOT_ID
                        .add "CITY NAME", HAF_CITY_NAME
                        .add "STATE CODE", HAF_STATE_CODE
                        .add "ZIP CODE", HAF_ZIP_CODE
                        .add "HOOKUP TYPE", HAF_HOOKUP_TYPE
                        .add "DWELLING TYPE", HAF_DWELLING_TYPE
                        .add "STATUS", HAF_STATUS
                        .add "SERVICEABILITY CODE", HAF_SERVICEABILITY_CODE
                        .add "INSTALLATION TYPE", HAF_INSTALLATION_TYPE
                        .add "NYS", HAF_NYS
                        .add "NYSBO", HAF_NYSBO
                        .add "NODE", HAF_NODE
                        .add "COMMENT", HAF_COMMENT
                        .add "HOUSE KEY", HAF_HOUSE_KEY
                        .add "AMP", HAF_AMP
                        .add "POWER SUPPLY", HAF_POWER_SUPPLY
                        .add "LAT", HAF_LAT
                        .add "LONG", HAF_LONG
                        .add "CENSUS BLOCK GROUP", HAF_CBG
                        .add "AWARD TYPE", HAF_AWARD_TYPE
                        .add "DROP LENGTH", HAF_DROP_LENGTH
                        .add "SIK READY", HAF_SIK_READY
                        .add "POLE AND PORT NUMBERS", HAF_POLE_PORT_NUMS
                End With
        End If
        Set haf_cols_dict = haf_cols_dict_p
End Property

Public Function haf_cols_from_string(ByVal haf_cols_text As String) As HAF_COLS
        haf_cols_from_string = enum_from_string(haf_cols_text, haf_cols_dict)
End Function

Public Function haf_cols_to_string(ByVal haf_cols_enum As HAF_COLS) As String
        haf_cols_to_string = enum_to_string(haf_cols_enum, haf_cols_dict)
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
        Set get_column_map = SCHEMA.get_column_map(sheet, HAF_HEADER_ROW, haf_cols_dict)
End Function

Public Function apply_conversion(ByVal col_name As HAF_COLS, ByVal in_value As String) As Variant
        Dim result As Variant
        result = Array(SCHEMA.apply_conversion(col_name, in_value, conversion_types))
        If IsObject(result(0)) Then
                Set apply_conversion = result(0)
        Else
                Let apply_conversion = result(0)
        End If
End Function
