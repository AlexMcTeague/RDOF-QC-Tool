VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TRCSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private super As SchemaSheet

Private Sub class_initialize()
        Set super = New SchemaSheet
End Sub

Public Property Set sheet(ByRef in_sheet As Worksheet)
        super.set_sheet "SCHEMA_TRC", in_sheet
End Property

Public Property Get sheet() As Worksheet
        Set sheet = super.sheet
End Property

Public Property Get Value(ByVal id As TRC_COLS, ByVal row_num As Long, Optional convert As Boolean = True) As Variant
Attribute Value.VB_UserMemId = 0
        Value = super(id, row_num, convert)
End Property

Public Function slice(ByRef cols As Variant, ByVal row_num As Long) As TRCSlice
        Dim result As TRCSlice
        Set result = New TRCSlice
        result.from_dict super.slice(cols, row_num)
        Set slice = result
End Function

Public Function get_letter(ByVal id As TRC_COLS) As String
        get_letter = super.get_letter(id)
End Function
