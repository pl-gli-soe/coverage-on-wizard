VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoverageIteration"
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
Public plt As String
Public pn As String
Public cbal As Long

Public resp As String
Public fup_code As String
Public supp_nm As String
Public duns As String
Public total_mrd_qty As String
Public del_conf As String
Public total_pus_qty As Long

Private slownik_cov_items As Dictionary
Private i As CoverageItem
Private j As CoordItem


Public Sub wypelnianie_arkusza_cov_rqmsami_i_pusami(ir As Range)
    
    Set i = Nothing
    Dim tmp As Range
    Set tmp = ir.Parent.Cells(1, ir.Column)
    
    'plt = ir.Parent.Cells(ir.Row, 1)
    'pn = ir.Parent.Cells(ir.Row, 2)
    
    Do
        ' ki = CStr(plt) & CStr(pn) & CStr(tmp)
        ki = CStr(tmp)
        
        
        Set i = Nothing
        
        On Error Resume Next
        Set i = slownik_cov_items(ki)
        
        
        If Not i Is Nothing Then
            ir = i.rqm
            ir.Offset(0, 1) = i.pus
        End If
    
        Set ir = ir.Offset(0, 3)
        Set tmp = tmp.Offset(0, 3)
    Loop Until Trim(tmp) = ""
End Sub



Public Sub sprawdz_weeki(ByRef najwczesniejszy_week As Variant, ByRef najpozniejszy_week As Variant)


    ' big side effect
    ' pobieramy na koniec wartosci na zewnatrz
    ' dla: najwczesniejszy_week i najpozniejszy_week

    For Each ki In slownik_cov_items.Keys
        Set i = slownik_cov_items(ki)
        
        If CStr(najwczesniejszy_week) = "" Then
            najwczesniejszy_week = i.ycw
        End If
        
        If CStr(najpozniejszy_week) = "" Then
            najpozniejszy_week = i.ycw
        End If
        
        
        najwczesniejszy_week = sprawdz_czy_jest_wczesniejszy(CStr(i.ycw), CStr(najwczesniejszy_week))
        najpozniejszy_week = sprawdz_czy_jest_pozniejszy(CStr(i.ycw), CStr(najpozniejszy_week))
    Next
End Sub

Private Function sprawdz_czy_jest_wczesniejszy(iycw As String, najwczes As String)
    sprawdz_czy_jest_wczesniejszy = CStr(najwczes)
    
    If iycw Like "*Y*CW*" Then
        If najwczes Like "*Y*CW*" Then
        
            ' sekcja sprawdzenia dwoch elementow ycw
            ' ===============================================
            
            
            
            If parsuj_y_cws_sprawdz_ktory_wiekszy(iycw, najwczes) = 2 Then
                sprawdz_czy_jest_wczesniejszy = CStr(iycw)
            End If
            ' ===============================================
        End If
    End If
End Function

Private Function sprawdz_czy_jest_pozniejszy(iycw As String, najpoz As String)
    sprawdz_czy_jest_pozniejszy = CStr(najpoz)
    
    If iycw Like "*Y*CW*" Then
        If najpoz Like "*Y*CW*" Then
        
            ' sekcja sprawdzenia dwoch elementow ycw
            ' ===============================================
            
            If parsuj_y_cws_sprawdz_ktory_wiekszy(iycw, najpoz) = 1 Then
                sprawdz_czy_jest_pozniejszy = CStr(iycw)
            End If
            ' ===============================================
        End If
    End If
End Function

Private Function parsuj_y_cws_sprawdz_ktory_wiekszy(ycw1, ycw2)
    ' wartosc mowi o tym ze pobieramy ycw2
    ' jesli parsuj_y_cws bedzie rowne 1 to wtedy pobieramy ycw1 ale to jest kod poza ta funkcja
    parsuj_y_cws_sprawdz_ktory_wiekszy = 0
    
    
    y1 = Mid(ycw1, 2, 4)
    y2 = Mid(ycw2, 2, 4)
    
    cw1 = Right(ycw1, 2)
    cw2 = Right(ycw2, 2)
    
    On Error Resume Next
    liczba_ycw1 = CLng(y1) * 100 + CLng(cw1)
    
    On Error Resume Next
    liczba_ycw2 = CLng(y2) * 100 + CLng(cw2)
    
    If liczba_ycw1 > liczba_ycw2 Then
        parsuj_y_cws_sprawdz_ktory_wiekszy = 1
    ElseIf liczba_ycw1 < liczba_ycw2 Then
        parsuj_y_cws_sprawdz_ktory_wiekszy = 2
    End If
End Function


Public Function get_cov_items() As Dictionary
    Set get_cov_items = slownik_cov_items
End Function


Public Function get_total_pus_qty() As Long
    
    get_total_pus_qty = CLng(total_pus_qty)
End Function

Public Sub calc_total_pus_qty_based_on_dictionary_coord_list(mj As CoordItem)


    total_pus_qty = 0
    
    For Each ki In slownik_cov_items.Keys
        Set mj = slownik_cov_items(ki)
        
        total_pus_qty = total_pus_qty + CLng(j.qty)
    Next
End Sub


Public Function get_puses() As Dictionary
    Set get_puses = slownik_cov_items
End Function


Private Sub Class_Initialize()
    Set slownik_cov_items = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set slownik_cov_items = Nothing
End Sub


Public Sub dodaj_lub_zmien_rqm_elementu_do_slownika_itemow(cwkey As String, r As Range)
    
    If slownik_cov_items.Exists(cwkey) Then
        
        Set i = slownik_cov_items(cwkey)
        
        If CStr(r.Offset(0, Cow.RQMS_QTY - Cow.RQMS_PLT)) = "" Then
            ' nooperation
        ElseIf IsNumeric(CStr(r.Offset(0, Cow.RQMS_QTY - Cow.RQMS_PLT))) Then
            i.rqm = CLng(i.rqm) + CLng(r.Offset(0, Cow.RQMS_QTY - Cow.RQMS_PLT))
        End If
    Else
        Set i = New CoverageItem
        i.ycw = CStr(cwkey)
        i.rqm = 0
        
        
        If IsNumeric(CStr(r.Offset(0, Cow.RQMS_QTY - Cow.RQMS_PLT))) Then
            i.rqm = CLng(r.Offset(0, Cow.RQMS_QTY - Cow.RQMS_PLT))
        End If
        
        slownik_cov_items.Add cwkey, i
    End If
End Sub


' ci.dodaj_lub_zmien_pus_elementu_do_slownika_itemow puses.Cells(r.Row, Cow.PUSES_DEL_DATE), r, puses_from_where
Public Sub dodaj_lub_zmien_pus_elementu_do_slownika_itemow(date_r As Range, r As Range, pfw As E_TYPE_OF_PUSES_FOR_COVERAGE)


    cwkey = ""
    cwkey = "Y" & CStr(Year(CDate(date_r))) & "CW"
    cw = CStr(Application.WorksheetFunction.IsoWeekNum(CDbl(CDate(date_r))))
    
    ' Debug.Assert cw <> "24"
    
    If Len(cw) = 1 Then
        cwkey = cwkey & "0" & cw
    ElseIf Len(cw) = 2 Then
        cwkey = cwkey & cw
    Else
        MsgBox "cos nie tak podczas proby dodania pusa do coveragea"
        End
    End If
    
    If cwkey Like "Y*CW*" Then
    
        
    
        Dim tmp As Long
        Dim scope_ref As Range
        Set scope_ref = r.Offset(0, Cow.PUSES_ON_MGO - Cow.PUSES_PLT)
        
        Dim czy_przegladamy_te_linie As Boolean
        czy_przegladamy_te_linie = False
        
        
        If pfw = E_TYPE_PUS_WIZARD Then
            
            ' --------------------------------------------------
            
            
            czy_przegladamy_te_linie = scope_ref.Offset(0, 1)
            
            
            ' --------------------------------------------------
        ElseIf pfw = E_TYPE_PUS_MGO Then
        
            ' --------------------------------------------------
            
            czy_przegladamy_te_linie = scope_ref
            
            ' --------------------------------------------------
        End If
        
        If czy_przegladamy_te_linie Then
        
        
            tmp = CLng(r.Offset(0, Cow.PUSES_QTY - Cow.PUSES_PLT))
    
            If slownik_cov_items.Exists(cwkey) Then
                
                Set i = slownik_cov_items(cwkey)
                i.pus = i.pus + CLng(tmp)
                            
            Else
                Set i = New CoverageItem
                i.ycw = CStr(cwkey)
                i.pus = CLng(tmp)
                
                'plt = r
                'pn = r.Offset(0, 1)
                
                
                ' slownik_cov_items.Add plt & pn & cwkey, i
                slownik_cov_items.Add cwkey, i
                
                
                
            End If
        End If
    End If
End Sub





' ci.dodaj_lub_zmien_pus_elementu_do_slownika_itemow puses.Cells(r.Row, Cow.PUSES_DEL_DATE), r, puses_from_where
Public Sub dodaj_lub_zmien_pus_dla_coord_list(pus_name As Range, r As Range, pfw As E_TYPE_OF_PUSES_FOR_COVERAGE)



    
    Dim tmp As Long
    Dim scope_ref As Range
    Set scope_ref = r.Offset(0, Cow.PUSES_ON_MGO - Cow.PUSES_PLT)
    
    Set j = Nothing
    
    Dim czy_przegladamy_te_linie As Boolean
    czy_przegladamy_te_linie = False
    
    
    If pfw = E_TYPE_PUS_WIZARD Then
        
        ' --------------------------------------------------
        
        
        czy_przegladamy_te_linie = scope_ref.Offset(0, 1)
        
        If czy_przegladamy_te_linie Then
            tmp = CLng(r.Offset(0, Cow.PUSES_QTY - Cow.PUSES_PLT))
        End If
        
        
        ' --------------------------------------------------
    ElseIf pfw = E_TYPE_PUS_MGO Then
    
        ' --------------------------------------------------
        
        czy_przegladamy_te_linie = scope_ref
        If czy_przegladamy_te_linie Then
            tmp = CLng(r.Offset(0, Cow.PUSES_DEL_QTY - Cow.PUSES_PLT))
        End If
        
        ' --------------------------------------------------
    End If
    
    If czy_przegladamy_te_linie Then

        ki = CStr(pus_name)
        
        If slownik_cov_items.Exists(ki) Then
            
            Set j = slownik_cov_items(ki)
            j.qty = CLng(tmp)
            j.del_date = CDate(r.Offset(0, Cow.PUSES_DEL_DATE - Cow.PUSES_PLT))
            j.pus_date = CDate(r.Offset(0, Cow.PUSES_PUS_DATE - Cow.PUSES_PLT))
            j.pus_name = CStr(r.Offset(0, Cow.PUSES_PUS_NAME - Cow.PUSES_PLT))
                        
        Else
            Set j = New CoordItem
            j.qty = CLng(tmp)
            j.del_date = CDate(r.Offset(0, Cow.PUSES_DEL_DATE - Cow.PUSES_PLT))
            j.pus_date = CDate(r.Offset(0, Cow.PUSES_PUS_DATE - Cow.PUSES_PLT))
            j.pus_name = CStr(r.Offset(0, Cow.PUSES_PUS_NAME - Cow.PUSES_PLT))
            
            slownik_cov_items.Add ki, j
            
            
            
        End If
    End If

End Sub
