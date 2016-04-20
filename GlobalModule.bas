Attribute VB_Name = "GlobalModule"
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

Global Const CURRENT_VERSION = "0.9"


'Public Enum E_RECV_TYPE
'    RECV_TBD = 0
'    ON_ZERO = 1
'    INLINE_WITH_QTY = 2
'    NOT_INLINE_WITH_QTY = 3
'    IN_TRANSIT = 4
'End Enum
' STR RECV SECTION
' =====================================================================
' =====================================================================
Global Const G_RECV_STR_TBD = "RECV TBD"
Global Const G_RECV_STR_ON_ZERO = "RECV NA ZERO"
Global Const G_RECV_STR_BOOKED = "BOOKED"
Global Const G_RECV_STR_BOOKED_NOT_INLINE = "BOOKED BUT NOT WITH SAME QTY"
Global Const G_RECV_STR_INTRANSIT = "IN TRANSIT"

' =====================================================================
' =====================================================================

Global Const STR_KROWA = "KROWA"
Global Const STR_COW = "COW"
Global Const G_STR_PTA = "PTA"
Global Const G_PODKRESLINIK_SEPARATOR = "_"
Global Const G_STR_CBAL = "CBAL"


' nazwy arkuszy wizarda
' ==========================================
Global Const MASTER_SH_NM = "MASTER"
Global Const PICKUPS_SH_NM = "PICKUPS"
Global Const DETAILS_SH_NM = "DETAILS"
Global Const COMMENT_SOURCE_SH_NM = "comment_source"
' ==========================================

' HISTORY SECTION
' ==========================================
Global Const ILE_DNI = 50
' ==========================================

' this workbook const sheets
' ==========================================
Global Const PUSES_SH_NM = "PUSes"
Global Const RQMS_SH_NM = "RQMs"
Global Const CBALS_SH_NM = "CBALs"
Global Const INPUT_SH_NM = "INPUT"
' ==========================================


Global Const LAST_ROW_IN_SH = 1048576


Global G_SOURCE_WIZARD As Workbook
Global G_FUP_CODE As String


Public Function try(s As String, proba As Variant) As String
    Dim sh As Worksheet
    Set sh = Nothing
    
    On Error Resume Next
    Set sh = ThisWorkbook.Sheets(s & "x" & CStr(proba))
    
    If Not sh Is Nothing Then
        arr = Split(s & "x" & CStr(proba), "x")
        tmp = CStr(try(CStr(arr(LBound(arr))), proba + 1))
        
        
        try = CStr(tmp)
    Else
        try = s & "x" & CStr(proba)
    End If
End Function

Public Function cowFirstRunout(l As Range) As String
    
    Do
        If l < 0 Then
            cowFirstRunout = CStr(l.Parent.Cells(1, l.Column - 2))
            Exit Function
        End If
        Set l = l.Offset(0, 3)
    Loop Until Trim(l) = ""
    
    ' to jest sztuczne
    cowFirstRunout = CStr(l.Parent.Cells(1, l.Column - 2 - 3))
End Function
