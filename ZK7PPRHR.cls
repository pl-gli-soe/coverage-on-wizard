VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZK7PPRHR"
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

Public Property Let sesja(ByRef arg As Object)
    Set sess0 = arg
End Property

Public Property Get sesja() As Object
    Set sesja = sess0
End Property

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
    KANBAN = Trim(sess0.screen.getString(5, 10, 4))
End Property

Public Property Let KANBAN(arg As String)
    sess0.screen.putString CStr(arg), 5, 10
End Property

Public Property Get ULOC() As String
    ULOC = Trim(sess0.screen.getString(4, 23, 8))
End Property

Public Property Let ULOC(arg As String)
    sess0.screen.putString CStr(arg), 4, 23
End Property

Public Function HRQMDATE(r As Integer, c As Integer) As String
    HRQMDATE = Trim(sess0.screen.getString(6 + 4 * r, 9 + 7 * c, 5))
End Function

Public Function HRQMTIME(r As Integer, c As Integer) As String
    HRQMTIME = Trim(sess0.screen.getString(7 + 4 * r, 9 + 7 * c, 5))
End Function

Public Function HQTY(r As Integer, c As Integer) As String
    HQTY = Trim(sess0.screen.getString(8 + 4 * r, 9 + 7 * c, 5))
End Function





