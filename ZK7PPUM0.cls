VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZK7PPUM0"
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

Public Property Get REC() As String
  REC = Trim(sess0.screen.getString(3, 65, 5))
End Property

Public Property Get OF() As String
  OF = Trim(sess0.screen.getString(3, 75, 5))
End Property



Public Function plt(i As Integer) As String
    plt = Trim(sess0.screen.getString(7 + 1 * i, 5, 2))
End Function

Public Function SUPPLIER(i As Integer) As String
    SUPPLIER = Trim(sess0.screen.getString(7 + 1 * i, 12, 9))
End Function
Public Function PU_DATE(i As Integer) As String
    PU_DATE = Trim(sess0.screen.getString(7 + 1 * i, 23, 8))
End Function

Public Function PUS_NBR(i As Integer) As String
    PUS_NBR = Trim(sess0.screen.getString(7 + 1 * i, 33, 9))
End Function

Public Function route(i As Integer) As String
    route = Trim(sess0.screen.getString(7 + 1 * i, 44, 6))
End Function

Public Function del_date(i As Integer) As String
    del_date = Trim(sess0.screen.getString(7 + 1 * i, 52, 8))
End Function

Public Function DEL_WINDW(i As Integer) As String
    DEL_WINDW = Trim(sess0.screen.getString(7 + 1 * i, 62, 9))
End Function

Public Function BEGIN_DEL(i As Integer) As String
    BEGIN_DEL = Trim(sess0.screen.getString(7 + 1 * i, 62, 4))
End Function
Public Function END_DEL(i As Integer) As String
    END_DEL = Trim(sess0.screen.getString(7 + 1 * i, 67, 4))
End Function

Public Function PART_QTY(i As Integer) As String
    PART_QTY = Trim(sess0.screen.getString(7 + 1 * i, 72, 9))
End Function
