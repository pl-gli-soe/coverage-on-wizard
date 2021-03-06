VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS9PO400"
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

Public Property Let sesja(ByRef arg As Object)
    Set sess0 = arg
End Property

Public Property Get sesja() As Object
    Set sesja = sess0
End Property

Public Property Get pn() As String
  pn = Trim(sess0.screen.getString(3, 19, 8))
End Property

Public Property Let pn(arg As String)
  sess0.screen.putString CStr(arg), 3, 19
End Property

Public Property Get plt() As String
  plt = Trim(sess0.screen.getString(3, 7, 2))
End Property

Public Property Let plt(arg As String)
  sess0.screen.putString CStr(arg), 3, 7
End Property

Public Property Get KANBAN() As String
  KANBAN = Trim(sess0.screen.getString(3, 35, 4))
End Property

Public Property Let KANBAN(arg As String)
  sess0.screen.putString CStr(arg), 3, 35
End Property



Public Function qty(i As Integer) As String
    qty = Trim(sess0.screen.getString(6 + 2 * i, 5, 8))
End Function
Public Function sdate(i As Integer) As String
    sdate = Trim(sess0.screen.getString(6 + 2 * i, 14, 6))
End Function
Public Function PCS(i As Integer) As String
    PCS = Trim(sess0.screen.getString(6 + 2 * i, 23, 3))
End Function
Public Function WGT(i As Integer) As String
    WGT = Trim(sess0.screen.getString(6 + 2 * i, 27, 6))
End Function
Public Function mode(i As Integer) As String
    mode = Trim(sess0.screen.getString(6 + 2 * i, 35, 2))
End Function
Public Function cd(i As Integer) As String
    cd = Trim(sess0.screen.getString(6 + 2 * i, 42, 1))
End Function
Public Function eda(i As Integer) As String
    eda = Trim(sess0.screen.getString(6 + 2 * i, 46, 6))
End Function
Public Function ETA(i As Integer) As String
    ETA = Trim(sess0.screen.getString(6 + 2 * i, 53, 4))
End Function
Public Function st(i As Integer) As String
    st = Trim(sess0.screen.getString(6 + 2 * i, 58, 1))
End Function
Public Function SID(i As Integer) As String
    SID = Trim(sess0.screen.getString(6 + 2 * i, 61, 9))
End Function
Public Function route(i As Integer) As String
    route = Trim(sess0.screen.getString(6 + 2 * i, 71, 7))
End Function
Public Function DOCK(i As Integer) As String
    DOCK = Trim(sess0.screen.getString(7 + 2 * i, 71, 7))
End Function
Public Function SHIP_CMNT(i As Integer) As String
    SHIP_CMNT = Trim(sess0.screen.getString(7 + 2 * i, 38, 32))
End Function
Public Function CARR(i As Integer) As String
    CARR = Trim(sess0.screen.getString(7 + 2 * i, 21, 16))
End Function
Public Function TRLR(i As Integer) As String
    TRLR = Trim(sess0.screen.getString(7 + 2 * i, 25, 11))
End Function
Public Function SUPPLIER(i As Integer) As String
    SUPPLIER = Trim(sess0.screen.getString(7 + 2 * i, 5, 15))
End Function
Public Function s(i As Integer) As String
    s = Trim(sess0.screen.getString(6 + 2 * i, 3, 1))
End Function

Public Sub C_SHIP_CMNT(str As String)
    sess0.screen.putString CStr(str), 7, 38
End Sub
