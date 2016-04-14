VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public pusname As String
Public pns As Collection

Public plt As String
Public duns As String
Public supp_nm As String

Public sdate As Date
Public eda As Date

Public fup_code As String


Public on_wizard As Boolean
Public on_mgo As Boolean

Public recv As E_RECV_TYPE

Public log As String


Private Sub Class_Initialize()
    Set pns = New Collection
End Sub

Private Sub Class_Terminate()
    Set pns = Nothing
End Sub
