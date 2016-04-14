VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FlatTablesHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private typeOfTable As E_TYPE_OF_FLAT_TABLE

Private ift As IFlatTable




Public Sub init(e1 As E_TYPE_OF_FLAT_TABLE)
    If e1 = E_FLAT_PUS Then
        Set ift = New FlatTablePus
    ElseIf e1 = E_FLAT_RQM Then
        Set ift = New FlatTableRqm
    End If
End Sub
