VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SchemaSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' All SCHEMA code credited to Brandon Trapp

Private col_map As Dictionary
Private sheet_p As Worksheet
Private schema_p As String

Public Sub set_sheet(ByVal schema_name As String, ByRef in_sheet As Worksheet)
    Set sheet_p = in_sheet
    Set col_map = Application.Run(schema_name & ".get_column_map", in_sheet)
    schema_p = schema_name
End Sub

Public Property Get sheet() As Worksheet
    Set sheet = sheet_p
End Property

Public Property Get Value(ByVal id As Variant, ByVal row_num As Long, Optional convert As Boolean = True) As Variant
Attribute Value.VB_UserMemId = 0
    Dim result As Variant
    result = Empty
    
    If col_map.Exists(id) Then
        result = sheet_p.Range(col_map(id) & row_num).Value
        
        If convert = True Then
            result = Application.Run(schema_p & ".apply_conversion", id, result)
        End If
    End If
    
    Value = result
End Property

Public Function slice(ByRef cols As Variant, ByVal row_num As Long) As Dictionary
    Dim result As Dictionary
    Set result = New Dictionary
    
    If TypeName(cols) = "Dictionary" Then
        For Each id In cols.Keys
            Dim cols_dict As Dictionary
            Set cols_dict = cols
            If result.Exists(id) Then
                result(id) = Value(id, row_num, cols_dict(id))
            Else
                result.add id, Value(id, row_num, cols_dict(id))
            End If
        Next
    ElseIf TypeName(cols) = "Variant()" Then
        For i = LBound(cols) To UBound(cols)
            id = cols(i)
            If result.Exists(id) Then
                result(id) = Value(id, row_num)
            Else
                result.add id, Value(id, row_num)
            End If
        Next
    Else
        Err.Raise vbObjectError + 513, "SchemaSheet", "Field 'cols' must be either a Scripting.Dictionary or Array object"
    End If
    
    Set slice = result
End Function

Public Function get_letter(ByVal id As Variant) As String
    Dim result As String
    result = ""
    
    If col_map.Exists(id) Then
        result = col_map(id)
    End If
    
    get_letter = result
End Function

' Function added by Alex McTeague
Public Function column(ByVal id As Variant) As Range
        Dim letter As String
        Dim result As Range
        
        If col_map.Exists(id) Then
            letter = col_map(id)
            Set result = sheet.Range(letter & 1).EntireColumn
        Else
            MsgBox "Error: Attempted to find a column that doesn't exist: " & ds_cols_to_string(id)
            Set result = Nothing
        End If
        
        Set column = result
End Function
