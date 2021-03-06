VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZK7PWRQM"
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

Public Property Get infocode() As String
    infocode = Trim(sess0.screen.getString(22, 2, 5))
End Property


Public Property Get plt() As String
    plt = Trim(sess0.screen.getString(3, 8, 2))
End Property

Public Property Let plt(arg As String)
    sess0.screen.putString CStr(arg), 3, 8
End Property

Public Property Get pn() As String
    pn = Trim(sess0.screen.getString(4, 11, 8))
End Property

Public Property Let pn(arg As String)
    sess0.screen.putString CStr(arg), 4, 11
End Property

Public Property Get KANBAN() As String
    KANBAN = Trim(sess0.screen.getString(5, 10, 4))
End Property

Public Property Let KANBAN(arg As String)
    sess0.screen.putString CStr(arg), 5, 10
End Property

Public Property Get PART_DESC() As String
    PART_DESC = Trim(sess0.screen.getString(4, 41, 10))
End Property



Public Function weeks(SCREEN_COLUMN As Integer) As String
    weeks = Trim(sess0.screen.getString(8, 22 + 8 * SCREEN_COLUMN, 7))
End Function

Public Function day_from_cw(cw As Integer, which_day As Integer) As Date
    
    Dim wn As Range, wd As Range, tf As Range
    Set wn = ThisWorkbook.Sheets("register").Range("weekNum")
    Set wd = ThisWorkbook.Sheets("register").Range("weekDay")
    Set tf = ThisWorkbook.Sheets("register").Range("todayFormula")
    ' swn - string on week num
    ' swd - string on week day
    Dim swn As String, swd As String, stf As String
    stf = Replace(tf.Formula, "=", "")
    swn = Replace(wn.Formula, "-1", stf)
    swd = Replace(wd.Formula, "-1", stf)
    
    current_cw = Evaluate(swn)
    current_weekday = Evaluate(swd)
    
    diff = cw - current_cw
    ' pamietaj ze jesli chodzi o formuly weekday to 1 - sunday, 2 - monday i tak dalej
    ' 8 dlatego, ze liczymy poniedzialek od 2 stad musimy kawalek dalej pojsc
    day_from_cw = Now + diff * 7 - 7 + (8 + which_day - current_weekday)
    
End Function

Public Function monday_from_cw(cw As Integer) As Date
    
    monday_from_cw = day_from_cw(cw, 1)
End Function

Public Function sunday_from_cw(cw As Integer) As Date
    
    sunday_from_cw = day_from_cw(cw, 7)
End Function

Public Function PLT_TOTAL(SCREEN_COLUMN As Integer, ByRef pf8_pressed As Boolean) As String
    
    ' zmienna pomocnicza dla ustawienia kolejnych element�w PLT TOTAL
    ' troche �le poniewa� za ka�dym razem gdy uruchamiam t� procedure
    ' wszystkie przeliczenia ciagle sie powtarzaja...
    Dim plttotalrow As Integer
    plttotalrow = fooplttotalrow(pf8_pressed)
    
    PLT_TOTAL = Trim(sess0.screen.getString(plttotalrow, 22 + 8 * SCREEN_COLUMN, 7))
    
    
End Function



Private Function fooplttotalrow(ByRef pf8_pressed As Boolean) As Integer

    ' While ((Trim(Sess0.screen.getString(22, 2, 5)) <> "R6101") And (Trim(Sess0.screen.getString(22, 2, 5)) <> "R6123"))
    While (Trim(sess0.screen.getString(22, 2, 5)) = "R6693") Or ((Trim(sess0.screen.getString(22, 2, 5)) = "R6102"))
        sess0.screen.sendKeys ("<pf8>")
        pf8_pressed = True
        Do
            DoEvents
        Loop Until sess0.screen.OIA.XStatus = 0
        
    Wend

    For x = 10 To 21
        If Trim(sess0.screen.getString(x, 8, 3)) Like "*TOT*" Then
            fooplttotalrow = x
            Exit For
        End If
    Next x
End Function

