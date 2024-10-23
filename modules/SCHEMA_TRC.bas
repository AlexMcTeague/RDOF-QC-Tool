Attribute VB_Name = "SCHEMA_TRC"
Public Const TRC_HEADER_ROW = 9

Enum TRC_COLS
        TRC_PATH_SPLIT
        TRC_ENC_UUID
        TRC_ENC_TYPE
        TRC_ENC_NAME
        TRC_ENC_LOCATION
        TRC_DEVICE_UUID_L
        TRC_DEVICE_NAME_L
        TRC_BUFFER_L
        TRC_FIBER_L
        TRC_PORT_NAME_L
        TRC_PORT_UUID_L
        TRC_SUB_CIRCUIT
        TRC_WAVELENGTH
        TRC_CIRCUIT
        TRC_SHEATH_FOOTAGE_L
        TRC_CONNECTION
        TRC_ATTENUATION
        TRC_CUMULATIVE_ATTENUATION
        TRC_PORT_UUID_R
        TRC_PORT_NAME_R
        TRC_BUFFER_R
        TRC_FIBER_R
        TRC_DEVICE_NAME_R
        TRC_DEVICE_UUID_R
        TRC_SHEATH_FOOTAGE_R
End Enum

Private conversion_types_p As Dictionary
Private trc_cols_dict_p As Dictionary

Private Property Get trc_cols_dict() As Dictionary
        If trc_cols_dict_p Is Nothing Then
                Set trc_cols_dict_p = New Dictionary
                With trc_cols_dict_p
                        .add "PATH SPLIT", TRC_PATH_SPLIT
                        .add "ENCLOSURE UUID", TRC_ENC_UUID
                        .add "ENCLOSURE TYPE", TRC_ENC_TYPE
                        .add "ENCLOSURE NAME", TRC_ENC_NAME
                        .add "ENCLOSURE LOCATION", TRC_ENC_LOCATION
                        .add "DEVICE UUID", Array(TRC_DEVICE_UUID_L, TRC_DEVICE_UUID_R)
                        .add "DEVICE NAME", Array(TRC_DEVICE_NAME_L, TRC_DEVICE_NAME_R)
                        .add "BUFFER", Array(TRC_BUFFER_L, TRC_BUFFER_R)
                        .add "FIBER", Array(TRC_FIBER_L, TRC_FIBER_R)
                        .add "PORT NAME", Array(TRC_PORT_NAME_L, TRC_PORT_NAME_R)
                        .add "PORT UUID", Array(TRC_PORT_UUID_L, TRC_PORT_UUID_R)
                        .add "SUB-CIRCUIT", TRC_SUB_CIRCUIT
                        .add "WAVELENGTH", TRC_WAVELENGTH
                        .add "CIRCUIT", TRC_CIRCUIT
                        .add "SHEATH FOOTAGE", Array(TRC_SHEATH_FOOTAGE_L, TRC_SHEATH_FOOTAGE_R)
                        .add "CONNECTION", TRC_CONNECTION
                        .add "ATTENUATION", TRC_ATTENUATION
                        .add "CUMULATIVE ATTENUATION", TRC_CUMULATIVE_ATTENUATION
                End With
        End If
        Set trc_cols_dict = trc_cols_dict_p
End Property

Public Function trc_cols_from_string(ByVal trc_cols_text As String) As TRC_COLS
        trc_cols_from_string = enum_from_string(trc_cols_text, trc_cols_dict)
End Function

Public Function trc_cols_to_string(ByVal trc_cols_enum As TRC_COLS) As String
        trc_cols_to_string = enum_to_string(trc_cols_enum, trc_cols_dict)
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
        Set get_column_map = SCHEMA.get_column_map(sheet, TRC_HEADER_ROW, trc_cols_dict)
End Function

Public Function apply_conversion(ByVal col_name As TRC_COLS, ByVal in_value As String) As Variant
        Dim result As Variant
        result = Array(SCHEMA.apply_conversion(col_name, in_value, conversion_types))
        If IsObject(result(0)) Then
                Set apply_conversion = result(0)
        Else
                Let apply_conversion = result(0)
        End If
End Function
