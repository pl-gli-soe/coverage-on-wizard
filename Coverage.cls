VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Coverage"
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


Implements ICoverage


Private rqms As Worksheet
Private puses As Worksheet
Private cbals As Worksheet

Private covsh As Worksheet
Private r As Range

Private slownik As Dictionary
Private ci As CoverageIteration


Private puses_from_where As E_TYPE_OF_PUSES_FOR_COVERAGE
Private stock_from_where As E_STOCK_FROM_WHERE



Private Sub ICoverage_makeCoverage(ich As InitConfigHandler)


    puses_from_where = ich.pusesForCoverage
    stock_from_where = ich.coverageStockBasedOnQuestion


    Set r = covsh.Range("a1")
    r = "COVERAGE based on: " & Cow.RQMS_SH_NM & ", " & Cow.PUSES_SH_NM & " & " & Cow.CBALS_SH_NM & ". Time: " & CStr(Now)
    
    Set r = r.Offset(1, 0)

    pobierz_cbale
    pobierz_rqmsy
    pobierz_pusy
    
    przygotuj_coverage_na_postawie_pobranych_danych
    koloruj
    dodaj_first_runout
    simpleAutoFit
End Sub

Private Sub dodaj_first_runout()
    Set r = covsh.Range("H3")
    
    Do
        r.FormulaR1C1 = "=cowFirstRunout(RC[5])"
        Set r = r.Offset(1, 0)
    Loop Until Trim(r.Offset(0, -1)) = ""
End Sub

Private Sub Class_Initialize()

    Set rqms = ThisWorkbook.Sheets(Cow.RQMS_SH_NM)
    Set puses = ThisWorkbook.Sheets(Cow.PUSES_SH_NM)
    Set cbals = ThisWorkbook.Sheets(Cow.CBALS_SH_NM)
    
    
    Set covsh = ThisWorkbook.Sheets.Add
    covsh.Name = try("COV_" & covsh.Name, 1)
    
    
    Set slownik = New Dictionary

End Sub

Private Sub Class_Terminate()
    Set slownik = Nothing
End Sub


Private Sub pobierz_cbale()

    Set r = cbals.Cells(2, 1)
    
    Do
        If Trim(r) <> "" Then
            
            If Not slownik.Exists(CStr(r) & CStr(r.Offset(0, 1))) Then
                
                Set ci = New CoverageIteration
                ci.plt = CStr(r.Offset(0, 0))
                ci.pn = CStr(r.Offset(0, 1))
                ci.cbal = CLng(r.Offset(0, Cow.CBAL_CBAL - 1))
                ci.del_conf = CStr(r.Offset(0, Cow.CBAL_DEL_CONF - 1))
                ci.duns = CStr(r.Offset(O, Cow.CBAL_DUNS - 1))
                ci.resp = CStr(r.Offset(0, Cow.CBAL_RESP - 1))
                ci.fup_code = CStr(r.Offset(0, Cow.CBAL_FUP_CODE - 1))
                ci.supp_nm = CStr(r.Offset(0, Cow.CBAL_SUPP_NM - 1))
                ci.total_mrd_qty = CStr(r.Offset(0, Cow.CBAL_TOTAL_MRD_QTY - 1))
                
                ' teraz sprawdz total pusy czy jest match
                ' cos tu nie hula
                ci.total_pus_qty = CLng(r.Offset(0, Cow.CBAL_TOTAL_PUS_QTY - 1))
                ' ci.calc_total_pus_qty_based_on_dictionary_coord_list New CoordItem
                'If CLng(ci.get_total_pus_qty) = CLng(r.Offset(0, Cow.CBAL_TOTAL_PUS_QTY)) Then
                '
                'End If
                
                
                
                slownik.Add CStr(r) & CStr(r.Offset(0, 1)), ci
            End If
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Sub

Private Sub pobierz_rqmsy()

    ' wszystkie dane zostaly pobrane nie bedziemy sie juz zajmowac wynajdywaniem kolejnym PNow
    ' skupimy sie na przegladaniu slownika i konfrontowaniu jego zwartosci z kolejnymi sideowymi arkuszami
    ' -----------------------------------------------------------------------------------------------------
    
    Set r = rqms.Cells(2, 1)
    
    Do
        If Trim(r) <> "" Then
            indx = CStr(r) & CStr(r.Offset(0, 1))
            If slownik.Exists(CStr(indx)) Then
                
                ' -----------------------------------------------------------------------------------------
                
                Set ci = slownik(CStr(indx))
                ci.dodaj_lub_zmien_rqm_elementu_do_slownika_itemow CStr(r.Offset(0, 3)), r
                
                ' -----------------------------------------------------------------------------------------
            End If
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    ' -----------------------------------------------------------------------------------------------------

End Sub

Private Sub pobierz_pusy()

    ' wszystkie dane zostaly pobrane nie bedziemy sie juz zajmowac wynajdywaniem kolejnym PNow
    ' skupimy sie na przegladaniu slownika i konfrontowaniu jego zwartosci z kolejnymi sideowymi arkuszami
    ' -----------------------------------------------------------------------------------------------------
    
    Set r = puses.Cells(2, 1)
    
    Do
        If Trim(r) <> "" Then
            ' indx = PLT & PN
            indx = CStr(r) & CStr(r.Offset(0, 1))
            If slownik.Exists(CStr(indx)) Then
                
                ' -----------------------------------------------------------------------------------------
                
                Set ci = slownik(CStr(indx))
                ci.dodaj_lub_zmien_pus_elementu_do_slownika_itemow puses.Cells(r.Row, Cow.PUSES_DEL_DATE), r, puses_from_where
                
                ' -----------------------------------------------------------------------------------------
            End If
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    ' -----------------------------------------------------------------------------------------------------
End Sub

Private Sub przygotuj_coverage_na_postawie_pobranych_danych()


    ' covsh - glowny koncowy raport
    ' slownik - przechowywane dane ze zrodel kazdy slownik ma key jako plt&pn
    ' a same values to elementy typu ci czyli coverage iteration
    
    Set r = covsh.Range("A2")
    r = "PLT"
    r.Offset(0, 1) = "PN"
    r.Offset(0, 2) = "DUNS"
    r.Offset(0, 3) = "SUPP NM"
    r.Offset(0, 4) = "RESP"
    r.Offset(0, 5) = "FUP CODE"
    r.Offset(0, 6) = "DEL CONF"
    r.Offset(0, 7) = "EST FIRST RUNOUT"
    r.Offset(0, 8) = "TOTAL MRD QTY"
    r.Offset(0, 9) = "CBAL"
    
    Set r = r.Offset(1, 0)
    For Each pltpn In slownik.Keys()
        ' iterujemy po ci
        ' =========================================================================
        Set ci = slownik(pltpn)
        r = ci.plt
        r.Offset(0, 1) = ci.pn
        r.Offset(0, 2) = ci.duns
        r.Offset(0, 3) = ci.supp_nm
        r.Offset(0, 4) = ci.resp
        r.Offset(0, 5) = ci.fup_code
        r.Offset(0, 6) = ci.del_conf
        ' 0,7 na runout
        r.Offset(0, 8) = ci.total_mrd_qty
        r.Offset(0, 9) = ci.cbal
        
        
        
        Set r = r.Offset(1, 0)
        
        ' =========================================================================
        
    Next
    
    Set r = covsh.Range("A2")
    
    
    przygotuj_labelki_dla_kolejnych_dni r.Offset(0, 10)
    
    Set r = covsh.Range("A3")
    
    Do
        Set ci = slownik(CStr(r) & CStr(r.Offset(0, 1)))
        
        ' iteracja - wypelnianie danych
        ci.wypelnianie_arkusza_cov_rqmsami_i_pusami r.Offset(0, 10)
        
        Set r = r.Offset(1, 0)
        ' jednak nie jest to potrzbne - wykonana zostala kopia na funkcji wypelnianie_arkusza_cov_rqmsami_i_pusami
        ' Set r = r.Parent.Cells(r.Row, 1).Offset(0, 9) ' czyli po krotce 10 kolumna ale zostawielem takie zagmatwane abym widzial ten sam schemat
    Loop Until Trim(r) = ""
    
    no_to_na_wielki_final_ebale stock_from_where
    

End Sub


Private Sub no_to_na_wielki_final_ebale(e As E_STOCK_FROM_WHERE)
    
    ' sekcja ebali
    ' ====================================================================
    ' ====================================================================
    ' ====================================================================
    
    Dim r As Range
    Set r = covsh.Range("M3")
    
    Do
    
        If e = E_STOCK_FROM_CBAL Then
            r.FormulaR1C1 = "=RC[-3] - RC[-2] + RC[-1]"
        ElseIf e = E_STOCK_FROM_TOTAL_MRD_QTY Then
            r.FormulaR1C1 = "=RC[-4] - RC[-2] + RC[-1]"
        End If
        
        
        Set r = r.Offset(0, Cow.COV_EBAL)
        
        If Trim(r.Parent.Cells(2, r.Column)) <> "" Then
            Do
                r.FormulaR1C1 = "=RC[-3] - RC[-2] + RC[-1]"
            
                Set r = r.Offset(0, Cow.COV_EBAL)
            Loop Until Trim(r.Parent.Cells(2, r.Column)) = ""
        End If
        
        Set r = r.Parent.Cells(r.Row, covsh.Range("M3").Column).Offset(1, 0)
    Loop Until Trim(r.Parent.Cells(r.Row, 2)) = ""
    
    
    ' ====================================================================
    ' ====================================================================
    ' ====================================================================
End Sub


Private Sub przygotuj_labelki_dla_kolejnych_dni(ByRef ir As Range)
    
    Dim najwczesniejszy_week As String, najpozniejszy_week As String
    
    najwczesniejszy_week = ""
    najpozniejszy_week = ""
    
    For Each pltpn In slownik.Keys()
        ' iterujemy po ci
        ' =========================================================================
        Set ci = slownik(pltpn)
        
        ci.sprawdz_weeki najwczesniejszy_week, najpozniejszy_week
        
        ' =========================================================================
        
    Next
    
    
    rysuj_labelki CStr(najwczesniejszy_week), CStr(najpozniejszy_week), ir
End Sub

Private Sub rysuj_labelki(w1 As String, w2 As String, ir As Range)

    
    i = 0
    w = w1
    w2_plus_1 = inkrementuj_o_jeden_week(CStr(w2))
    Do
        ir.Offset(-1, i * 3) = w
        ir.Offset(0, i * 3) = "RQM"
        ir.Offset(0, i * 3 + 1) = "IT"
        ir.Offset(0, i * 3 + 2) = "EBAL"
        w = inkrementuj_o_jeden_week(CStr(w))
        i = i + 1
        
        
        ' taka heurystyczna granica gdyby cos z weekami bylo nie tak
        If i > 100 Then
            Exit Do
        End If
    Loop Until CStr(w) = CStr(w2_plus_1)
End Sub


Private Function inkrementuj_o_jeden_week(iw As String)

    y = CLng(Mid(iw, 2, 4))
    ' Application.WorksheetFunction.Weekday(Date, 3) pn jako zero!
    najblizszy_poniedzialek = CDate(Date - Application.WorksheetFunction.Weekday(Date, 3))
    
    cw = CLng(Right(iw, 2))
    
    Do
        yp = Year(najblizszy_poniedzialek)
        
        ycw = y * 100 + CLng(cw)
        ycw_np = yp * 100 + CLng(Application.WorksheetFunction.IsoWeekNum(CDbl(najblizszy_poniedzialek)))
        
        If ycw < ycw_np Then
            najblizszy_poniedzialek = CDate(najblizszy_poniedzialek - 7)
        ElseIf ycw > ycw_np Then
            najblizszy_poniedzialek = CDate(najblizszy_poniedzialek + 7)
        End If
        
    Loop Until CLng(ycw) = CLng(ycw_np)
    
    przesuniety_tydzien = CDate(najblizszy_poniedzialek + 7)
    
    
    cstr_cw = CStr(Application.WorksheetFunction.IsoWeekNum(przesuniety_tydzien))
    
    If Len(cstr_cw) = 1 Then
        cstr_cw = "0" & cstr_cw
    End If
    
    inkrementuj_o_jeden_week = "Y" & CStr(Year(przesuniety_tydzien) & "CW" & CStr(cstr_cw))
End Function

Private Sub koloruj()

    
    koloruj_labelki
    minusy
    znajdz_del_conf_dla_linii
    grid_na_cale
    ramki_na_ebale_plus_boldy_na_wybranych_kolumnach
End Sub


Private Sub koloruj_labelki()


    Set r = covsh.Range("A1")
    
    If Trim(r.Offset(1, 0)) <> "" Then
        lr = r.End(xlDown).Row
    Else
        lr = 3
    End If
    
    
    If Trim(r.Offset(1, 1)) <> "" Then
        lc = r.Offset(1, 0).End(xlToRight).Column
    Else
        lc = 2
    End If
    
    Dim l As TheLayout
    Set l = New TheLayout
    
    'Set r = covsh.Range(covsh.Cells(1, 1), covsh.Cells(1, lc))
    'l.fillSolidGridLines r, RGB(0, 0, 250)
    
    Set r = covsh.Range(covsh.Cells(2, 1), covsh.Cells(2, lc))
    l.fillSolidGridLines r, RGB(0, 0, 0)

End Sub

Private Sub minusy()


    Set r = covsh.Range("A2")
    
    If Trim(r.Offset(1, 0)) <> "" Then
        lr = r.End(xlDown).Row
    Else
        lr = 3
    End If
    
    
    If Trim(r.Offset(0, 1)) <> "" Then
        lc = r.End(xlToRight).Column
    Else
        lc = 2
    End If
    
    
    Set r = covsh.Range(covsh.Cells(3, 1), covsh.Cells(lr, lc))
    r.NumberFormat = "0_ ;[Red]-0 "
End Sub

Private Sub znajdz_del_conf_dla_linii()

    Dim rqm As Range
    Dim tmp As Range
    Dim l As TheLayout
    
    
    Set l = New TheLayout

    Set r = covsh.Range("G3")
    Set rqm = covsh.Range("K3")
    
    Do
        Do
            If CStr(r) = CStr(rqm.Parent.Cells(1, rqm.Column)) Then
                Set tmp = rqm.Parent.Range(rqm, rqm.Offset(0, 2))
                l.backColor tmp, RGB(0, 160, 100)
                Exit Do
            End If
            
            Set rqm = rqm.Offset(0, 3)
        Loop Until Trim(CStr(rqm.Parent.Cells(1, rqm.Column))) = ""
        
        Set rqm = rqm.Parent.Cells(r.Row, covsh.Range("K3").Column).Offset(1, 0)
        
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Sub

Private Sub ramki_na_ebale_plus_boldy_na_wybranych_kolumnach()

    Set r = covsh.Range("A2")
    
    If Trim(r.Offset(1, 0)) <> "" Then
        lr = r.End(xlDown).Row
    Else
        lr = 3
    End If
    
    
    If Trim(r.Offset(0, 1)) <> "" Then
        lc = r.End(xlToRight).Column
    Else
        lc = 2
    End If
    
    Dim t As TheLayout
    Set t = New TheLayout
    
    Do
        If r = "EBAL" Then
            t.fillSolidGridLines covsh.Range(covsh.Cells(2, r.Column), covsh.Cells(lr, r.Column)), RGB(10, 40, 10)
            covsh.Range(covsh.Cells(2, r.Column), covsh.Cells(lr, r.Column)).Font.Bold = True
        ElseIf r = "DEL CONF" Or r = "PLT" Or r = "PN" Then
            covsh.Range(covsh.Cells(2, r.Column), covsh.Cells(lr, r.Column)).Font.Bold = True
        End If
        Set r = r.Offset(0, 1)
    Loop Until Trim(r) = ""
End Sub

Private Sub grid_na_cale()

    Set r = covsh.Range("A2")
    
    If Trim(r.Offset(1, 0)) <> "" Then
        lr = r.End(xlDown).Row
    Else
        lr = 3
    End If
    
    
    If Trim(r.Offset(0, 1)) <> "" Then
        lc = r.End(xlToRight).Column
    Else
        lc = 2
    End If
    
    Dim t As TheLayout
    Set t = New TheLayout
    
    t.fillThinGridLines Range(covsh.Cells(2, 1), covsh.Cells(lr, lc)), RGB(10, 10, 10)
End Sub


Private Sub simpleAutoFit()

    With covsh
        
        .Columns("B:J").AutoFit
    End With
End Sub
