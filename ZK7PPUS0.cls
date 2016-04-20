VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZK7PPUS0"
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
  pn = Trim(sess0.screen.getString(11, 27, 8))
End Property

Public Property Let pn(arg As String)
  sess0.screen.putString CStr(arg), 11, 27
End Property

Public Property Get KANBAN_NUMBER() As String
  KANBAN_NUMBER = Trim(sess0.screen.getString(13, 27, 4))
End Property

Public Property Let KANBAN_NUMBER(arg As String)
  sess0.screen.putString CStr(arg), 13, 27
End Property

Public Property Get DELIVERY_DATE_FROM() As String
  DELIVERY_DATE_FROM = Trim(sess0.screen.getString(15, 35, 10))
End Property

Public Property Let DELIVERY_DATE_FROM(arg As String)
  sess0.screen.putString CStr(arg), 15, 35
End Property

Public Property Get RDC_DATE_FROM() As String
  RDC_DATE_FROM = Trim(sess0.screen.getString(17, 35, 10))
End Property

Public Property Let RDC_DATE_FROM(arg As String)
  sess0.screen.putString CStr(arg), 17, 35
End Property

Public Property Get PICKUP_DATE_FROM() As String
  PICKUP_DATE_FROM = Trim(sess0.screen.getString(19, 35, 10))
End Property

Public Property Let PICKUP_DATE_FROM(arg As String)
  sess0.screen.putString CStr(arg), 19, 35
End Property

Public Property Get DELIVERY_DATE_TO() As String
  DELIVERY_DATE_TO = Trim(sess0.screen.getString(15, 54, 10))
End Property

Public Property Let DELIVERY_DATE_TO(arg As String)
  sess0.screen.putString CStr(arg), 15, 54
End Property

Public Property Get RDC_DATE_TO() As String
  RDC_DATE_TO = Trim(sess0.screen.getString(17, 54, 10))
End Property

Public Property Let RDC_DATE_TO(arg As String)
  sess0.screen.putString CStr(arg), 17, 54
End Property

Public Property Get PICKUP_DATE_TO() As String
  PICKUP_DATE_TO = Trim(sess0.screen.getString(19, 54, 10))
End Property

Public Property Let PICKUP_DATE_TO(arg As String)
  sess0.screen.putString CStr(arg), 19, 54
End Property

Public Property Get COUNTRY_CODE() As String
  COUNTRY_CODE = Trim(sess0.screen.getString(7, 67, 3))
End Property

Public Property Get STATE_PROVINCE() As String
  STATE_PROVINCE = Trim(sess0.screen.getString(9, 67, 2))
End Property

Public Property Get PICKUP_CARRIER() As String
  PICKUP_CARRIER = Trim(sess0.screen.getString(11, 67, 4))
End Property

Public Property Get RDC() As String
  RDC = Trim(sess0.screen.getString(13, 67, 4))
End Property

Public Property Get fu() As String
  fu = Trim(sess0.screen.getString(4, 36, 2))
End Property

Public Property Get ROUTE_SUFFIX() As String
  ROUTE_SUFFIX = Trim(sess0.screen.getString(7, 27, 3))
End Property

Public Property Get SUPPLIER_ALIAS() As String
  SUPPLIER_ALIAS = Trim(sess0.screen.getString(9, 27, 9))
End Property
