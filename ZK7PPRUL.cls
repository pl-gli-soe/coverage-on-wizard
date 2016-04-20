VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZK7PPRUL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2015 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
Private sess0 As Object
' zmienna przechowujaca reszte danych z zk7pdrqm
Private ULOC As Collection

Public Property Let sesja(ByRef arg As Object)
    Set sess0 = arg
End Property

Public Property Get sesja() As Object
    Set sesja = sess0
End Property


Public Property Let p_uloc(ByRef arg As Object)
    Set ULOC = arg
End Property

Public Property Get p_uloc() As Object
    Set p_uloc = ULOC
End Property

Private Sub Class_Initialize()
    Set ULOC = New Collection
End Sub

Private Sub Class_Terminate()
    Set ULOC = Nothing
End Sub




Public Property Get plt() As String
    plt = Trim(sess0.screen.getString(3, 8, 2))
End Property

Public Property Let plt(arg As String)
    sess0.screen.putString CStr(arg), 3, 8
End Property

Public Property Get pn() As String
    pn = Trim(sess0.screen.getString(4, 8, 8))
End Property

Public Property Let pn(arg As String)
    sess0.screen.putString CStr(arg), 4, 8
End Property

Public Property Get KANBAN() As String
    KANBAN = Trim(sess0.screen.getString(4, 26, 4))
End Property

Public Property Let KANBAN(arg As String)
    sess0.screen.putString CStr(arg), 4, 26
End Property

Public Sub fill_uloc_collection(i_mgo As MGO)

    Do
        For x = 8 To 21
            If (Trim(sess0.screen.getString(x, 6, 4)) <> "MISC") And (sess0.screen.getString(x, 6, 4) <> "    ") Then
                ULOC.Add Trim(sess0.screen.getString(x, 6, 8))
            End If
        Next x
        
        If sess0.screen.getString(22, 2, 5) = "R6086" Then
            Exit Do
        End If
        
        If sess0.screen.getString(22, 2, 5) = "R6017" Then
            Exit Do
        End If
        
        If sess0.screen.getString(22, 2, 5) = "R6102" Then
            i_mgo.sendKeys "<pf8>"
            i_mgo.waitForMgo
        End If
        
    Loop While True
    
    ' MsgBox ULOC.COUNT
End Sub





