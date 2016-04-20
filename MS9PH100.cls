VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS9PH100"
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

Public Property Get plt() As String
  plt = Trim(sess0.screen.getString(4, 6, 2))
End Property

Public Property Let plt(arg As String)
  sess0.screen.putString CStr(arg), 4, 6
End Property

Public Property Get pn() As String
  pn = Trim(sess0.screen.getString(4, 17, 8))
End Property

Public Property Let pn(arg As String)
  sess0.screen.putString "        ", 4, 17
  sess0.screen.putString CStr(arg), 4, 17
End Property

Public Property Get DATE_FROM() As String
  DATE_FROM = Trim(sess0.screen.getString(6, 8, 8))
End Property

Public Property Let DATE_FROM(arg As String)
  sess0.screen.putString CStr(arg), 6, 8
End Property

Public Function fDTYP() As String
  fDTYP = Trim(sess0.screen.getString(7, 8, 11))
End Function

Public Sub setDTYP(arg1 As String, arg2 As String)
  sess0.screen.putString CStr(arg1), 7, 8
  sess0.screen.putString CStr(arg2), 7, 15
End Sub

Public Function SID(i As Integer) As String
    SID = Trim(sess0.screen.getString(11 + 1 * i, 55, 9))
End Function

Public Function DTYP(i As Integer) As String
    DTYP = Trim(sess0.screen.getString(11 + 1 * i, 17, 4))
End Function

Public Function RECV_DATE(i As Integer) As String
    RECV_DATE = Trim(sess0.screen.getString(11 + 1 * i, 30, 4))
End Function

Public Function duns(i As Integer) As String
    duns = Trim(sess0.screen.getString(11 + 1 * i, 65, 9))
End Function

Public Function qty(i As Integer) As String
    qty = Trim(sess0.screen.getString(11 + 1 * i, 35, 10))
End Function

Public Property Get infocode() As String
    infocode = Trim(sess0.screen.getString(22, 2, 5))
End Property
