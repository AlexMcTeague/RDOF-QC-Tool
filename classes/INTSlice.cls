VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INTSlice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private data_p As Dictionary

Private Sub class_initialize()
        Set data_p = New Dictionary
End Sub

Public Property Get Value(ByVal id As BOM_INTERNAL_COLS) As Variant
Attribute Value.VB_UserMemId = 0
        Dim result As Variant
                result = Empty
        If data_p.Exists(id) Then
                result = data_p(id)
        End If

        Value = result
End Property

Public Property Let Value(ByVal id As BOM_INTERNAL_COLS, ByVal new_value As Variant)
Attribute Value.VB_UserMemId = 0
        If data_p.Exists(id) Then
                data_p(id) = new_value
        Else
                data_p.add id, new_value
        End If
End Property

Public Sub from_dict(ByRef in_dict As Dictionary)
        Set data_p = in_dict
End Sub
