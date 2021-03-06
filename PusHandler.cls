VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PusHandler"
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
Private puses As Dictionary
Private p As pus
Private item As PnInPus


Private this_psh As Worksheet

Private ki As String
Private ki1 As String
Private ki2 As String


Public Sub refresh_puses_sheet()

    Set this_psh = ThisWorkbook.Sheets(Cow.PUSES_SH_NM)
End Sub


Public Function get_puses() As Dictionary
    Set get_puses = puses
End Function

' extra system / mgo section
' ===================================================================
' ===================================================================
' ===================================================================
' 1.
Public Sub take_puses_from_mgo(m As MGO, scr As ZK7PPUM0, strpn As String)

    ' downloading data section
    ' ==================================================
    
    If m.actualScreen() = "ZK7PPUM0" Then
        ' ----------------------------------------------
        If CStr(m.getCode("zk7ppum0")) = "R6086" Then
            ' std jedna strona
            ' ==========================================
            
            For x = 0 To 12
            
                ki = CStr(scr.PUS_NBR(Int(x)))
                
                If Trim(ki) <> "" Then
                    If puses.Exists(ki) Then
                    
                        Set p = puses(ki)
                        
                        ' sprawdz pn
                        Set item = Nothing
                        Dim iitem As PnInPus
                        For Each iitem In p.pns
                            If Trim(CStr(iitem.pn)) = strpn Then
                                Set item = iitem
                                Exit For
                            End If
                        Next iitem
                        
                        If Not item Is Nothing Then
                            item.del_qty = scr.PART_QTY(Int(x))
                        Else
                            
                            Set item = Nothing
                            Set item = New PnInPus
                            
                            item.pn = strpn
                            item.qty = 0
                            item.del_qty = scr.PART_QTY(Int(x))
                            item.recv = 0
                            
                            item.content_source = CONTENT_FROM_PUS
                            
                            p.pns.Add item
                            p.on_wizard = False
                        End If
                        p.on_mgo = True
                    
                    Else
                    
                    
                        Set item = Nothing
                        Set item = New PnInPus
                        
                        item.pn = strpn
                        item.qty = 0
                        item.del_qty = scr.PART_QTY(Int(x))
                        item.recv = 0
                        
                        item.content_source = CONTENT_FROM_PUS
                        
                        
                        Set p = Nothing
                        Set p = New pus
                        
                        With p
                        
                            .plt = scr.plt(Int(x))
                            .duns = scr.SUPPLIER(Int(x))
                            .supp_nm = ""
                        
                            .eda = CDate(m.convertToDateFromZK7PPUM0Date(Trim(scr.del_date(Int(x)))))
                            .pusname = CStr(ki)
                            .recv = RECV_TBD
                            
                            .on_mgo = True
                            .on_wizard = False
                            
                            .fup_code = ""
                            
                            .sdate = CDate(m.convertToDateFromZK7PPUM0Date(Trim(scr.PU_DATE(Int(x)))))
                            .pns.Add item
                        End With
                        
                        puses.Add CStr(ki), p
                    End If
                End If
            Next x
            
            
            ' ==========================================
        End If
        
        ' ----------------------------------------------
    End If
    
    ' ==================================================
End Sub

' 1.
Public Sub take_asns_from_mgo(m As MGO, scr As MS9PO400)

    ' downloading data section
    ' ==================================================
    
    If CStr(m.actualScreen()) = "MS9PO400" Then
    
    If CStr(m.getCode("ms9po400")) <> "I6155" Then
        
        ' ----------------------------------------------
        Do
            For i = 0 To 7
            
                ' scr.qty (i)
                ki1 = CStr(scr.SID(Int(i)))
                
                If ki1 <> "" Then
                    'asn_comment = CStr(scr.SHIP_CMNT(Int(i)))
                    'ki2 = match_keys_with_asn_comment(CStr(asn_comment))
                    
                    ' still nothing then try to find pta numer in cmnt
                    ' i to na samym koncu
                    
                    ' sekcja PTA
                    ' ------------------------------------------
                    'If ki2 = "" Then
                    '    ki2 = try_to_find_pta_number_in_comment(CStr(asn_comment))
                    'End If
                    ' ------------------------------------------
                    
    
                    'If ki2 = "" Then
                    '    ki = ki1
                    'Else
                    '    ki = ki2
                    'End If
                    
                    ' olewam komentarz skupiam sie tylko na param SID.
                    ki = ki1
                    
                    
                    
                    If puses.Exists(ki) Then
                        
                        
                        
                        Set p = puses(ki)
                        
                        ' sprawdz pn
                        Set item = Nothing
                        Dim iitem As PnInPus
                        For Each iitem In p.pns
                            If Trim(CStr(iitem.pn)) = scr.pn Then
                                Set item = iitem
                                Exit For
                            End If
                        Next iitem
                        
                        If Not item Is Nothing Then
                            item.del_qty = scr.qty(Int(i))
                        Else
                            
                            Set item = Nothing
                            Set item = New PnInPus
                            
                            item.pn = scr.pn
                            item.qty = 0
                            item.del_qty = scr.qty(Int(i))
                            item.recv = 0
                            item.content_source = CONTENT_FROM_ASN
                            
                            p.pns.Add item
                            p.on_wizard = False
                        End If
                        p.on_mgo = True
                    Else
                    
                        Set item = Nothing
                        Set item = New PnInPus
                        
                        item.pn = scr.pn
                        item.qty = 0
                        item.del_qty = scr.qty(Int(i))
                        item.recv = 0
                        
                        item.content_source = CONTENT_FROM_ASN
                        
                        
                        Set p = Nothing
                        Set p = New pus
                        
                        With p
                        
                            .plt = scr.plt
                            .duns = ""
                            .supp_nm = ""
                        
                            .eda = CDate(m.convertToDateFromMS9PO400Date(scr.eda(Int(i))))
                            .pusname = CStr(ki)
                            .recv = RECV_TBD
                            
                            .on_mgo = True
                            .on_wizard = False
                            
                            .fup_code = ""
                            
                            .sdate = CDate(m.convertToDateFromMS9PO400Date(scr.sdate(Int(i))))
                            .pns.Add item
                        End With
                        
                        puses.Add CStr(ki), p
                    End If
                End If
            Next i
            
            If CStr(m.getCode("ms9po400")) = "I5487" Then
                Exit Do
            End If
            
            If CStr(m.getCode("ms9po400")) = "I7147" Then
                Exit Do
            End If
            
            If CStr(m.getCode("ms9po400")) = "I4300" Then
                m.sendKeys "<pf8>"
            End If
            
            
        Loop While True
        
        
        ' ----------------------------------------------
    End If
    
    End If
    
    ' ==================================================

End Sub


Private Function try_to_find_pta_number_in_comment(asn_comment As String) As String
    try_to_find_pta_number_in_comment = ""
    
    If CStr(asn_comment) Like "*" & Cow.G_STR_PTA & "*" Then
        
        tmp_arr = Split(asn_comment, Cow.G_PODKRESLINIK_SEPARATOR)
        
        For x = LBound(tmp_arr) To UBound(tmp_arr)
            If tmp_arr(x) Like "*" & Cow.G_STR_PTA & "*" Then
            
                ' to jest rozwiazanie statyczne na nastepne 800 tys pta.
            
                'od_prawej = Right(CStr(tmp_arr(x)), Len(CStr(tmp_arr(x))) - 3)
                'try_to_find_pta_number_in_comment = Left(CStr(tmp_arr(x)), 3) & "000" & CStr(od_prawej)
                
                try_to_find_pta_number_in_comment = Replace(CStr(tmp_arr(x)), "0", "")
                
                Exit Function
            End If
        Next x
    End If
End Function

Private Function match_keys_with_asn_comment(asn_comment As String) As String
    match_keys_with_asn_comment = ""
    
    For Each klucz In puses.Keys()
        If asn_comment Like "*" & CStr(klucz) & "*" Then
            match_keys_with_asn_comment = CStr(klucz)
            Exit Function
        End If
    Next
End Function

' 2.
Public Sub merge_puses_into_this_puses_sh()

    Set this_psh = ThisWorkbook.Sheets(Cow.PUSES_SH_NM)
    
    ' no i tak lecimy caly arkusz i porownanie na pelnej
    ' MATCHING SECTION FOR PUSes sheet
    ' =======================================================
    
    
    'x = CLng(first_empty_row(this_psh.Cells(1, 1)))
    'For Each strki In puses.Keys()
    '    ' i gitara
    '    ' =================================================
    '
    '    Set p = puses(strki)
    '    For Each item In p.pns
    '
    '        this_psh.Cells(x, 1) = p.plt
    '        this_psh.Cells(x, 2) = item.pn
    '        this_psh.Cells(x, 3) = p.fup_code
    '        this_psh.Cells(x, 4) = p.sdate
    '        this_psh.Cells(x, 5) = p.eda
    '        this_psh.Cells(x, 6) = item.qty
    '        this_psh.Cells(x, 7) = item.del_qty
    '        this_psh.Cells(x, 8) = item.recv
    '        this_psh.Cells(x, 9) = p.recv
    '        this_psh.Cells(x, 10) = p.pusname
    '        this_psh.Cells(x, 11) = p.duns
    '        this_psh.Cells(x, 12) = p.supp_nm
    '        this_psh.Cells(x, 13) = p.on_mgo
    '        this_psh.Cells(x, 14) = p.on_wizard
    '        x = x + 1
    '    Next item
    '
    '    ' =================================================
    'Next
    
    ' x = CLng(first_empty_row(this_psh.Cells(1, 1)))
    For Each strki In puses.Keys()
        ' =================================================
        
        Set p = puses(strki)
        
        ' sprawdz teraz ten item czy nie ma go juz na arkuszu
        pracuj_na_danych
        
        ' =================================================
    Next
    
    ' =======================================================
End Sub

Private Sub pracuj_na_danych()


    Set this_psh = ThisWorkbook.Sheets(Cow.PUSES_SH_NM)

    x = 2
    If Trim(this_psh.Cells(x, 1)) <> "" Then
        ' to znaczy ze jest na czym pracowac w ogole
        ' ------------------------------------------
        
        
            Dim czy_dopisujemy_nowy_rekord_na_dnie As Boolean
            
            If p.pns.COUNT > 0 Then
                On Error Resume Next
                For Each item In p.pns
                    
                    
                    If Not item Is Nothing Then
                    
                    
                    
                    
                        x = 2
                        czy_dopisujemy_nowy_rekord_na_dnie = True
                        Do
                            Dim r As Range
                            Set r = this_psh.Cells(x, 1)
                            
                            
                            ' part numbery sa posrotwane wiec swobodnie w pierwszej kolejnosci sprawdzamy
                            ' czy dalej jestesmy w jednym worku
                            If CLng(item.pn) = CLng(this_psh.Cells(x, Cow.PUSES_PN)) Then
                            
                                'from_cell = Replace(CStr(this_psh.Cells(x, Cow.PUSES_PUS_NAME)), "0", "")
                                'from_pus_name_from_object = Replace(CStr(p.pusname), "0", "")
                                'If (from_pus_name_from_object = from_cell) Or takie_same_4_od_konca(from_cell, from_pus_name_from_object) Then
                                
                                from_cell = Trim(CStr(this_psh.Cells(x, Cow.PUSES_PUS_NAME)))
                                from_pus_name_from_object = Trim(CStr(p.pusname))
                                
                                ' tutaj bardzo restrykcyjnie sprawdzmy czy nazwy pusow sa identyczne
                                If Trim(from_pus_name_from_object) = Trim(from_cell) Then
                                
                                
                                    ' this_psh.Cells(x, Cow.PUSES_PUS_NAME) = from_cell
                                
                                    this_psh.Cells(x, Cow.PUSES_DEL_QTY) = item.del_qty
                                    this_psh.Cells(x, Cow.PUSES_ON_MGO) = p.on_mgo
                                    
                                    ' recv
                                    this_psh.Cells(x, Cow.PUSES_RECV) = item.recv
                                    
                                    ' Debug.Print item.content_source
                                    
                                    
                                    def_recv_status this_psh.Cells(x, Cow.PUSES_BOOL_RECV), _
                                        p, item, CLng(this_psh.Cells(x, Cow.PUSES_RECV)), _
                                        CLng(this_psh.Cells(x, Cow.PUSES_QTY)), _
                                        CLng(this_psh.Cells(x, Cow.PUSES_DEL_QTY)), _
                                        True
                                    
                                    ' checks
                                    p.log = ""
                                    
                                    ' check dates
                                    ' ----------------------------------------------------------------------------
                                    'If CDbl(CDate(this_psh.Cells(x, Cow.PUSES_DEL_DATE)) - CDate(p.eda)) > 1# Then
                                    '    p.log = " wizard & mgo data not inline"
                                    'End If
                                    
                                    'If CDbl(CDate(this_psh.Cells(x, Cow.PUSES_PUS_DATE)) - CDate(p.sdate)) > 1# Then
                                    '    p.log = " wizard & mgo data not inline"
                                    'End If
                                    ' ----------------------------------------------------------------------------
                                    
                                    ' check qty
                                    ' ----------------------------------------------------------------------------
                                    'If Math.Abs(CDbl(this_psh.Cells(x, Cow.PUSES_QTY)) - CDbl(item.del_qty)) > 0# Then
                                    '    p.log = " wizard & mgo data not inline"
                                    'End If
                                    ' ----------------------------------------------------------------------------
                                    
                                    'this_psh.Cells(x, Cow.PUSES_LOG) = p.log
                                
                                    ' Exit For
                                    czy_dopisujemy_nowy_rekord_na_dnie = False
                                End If
                            End If
                            
                            x = x + 1
                            Set r = r.Offset(1, 0)
                        Loop Until Trim(r) = ""
                        
                        ' nie znalazl zatem:
                        
                        
                        If czy_dopisujemy_nowy_rekord_na_dnie Then
                            x = r.Row
                        
                            this_psh.Cells(x, 1) = p.plt
                            this_psh.Cells(x, 2) = item.pn
                            this_psh.Cells(x, 3) = p.fup_code
                            this_psh.Cells(x, 4) = p.sdate
                            this_psh.Cells(x, 5) = p.eda
                            this_psh.Cells(x, 6) = item.qty
                            this_psh.Cells(x, 7) = item.del_qty
                            this_psh.Cells(x, 8) = item.recv
                            
                            ' def_recv_status this_psh.Cells(x, 9), p, item, CLng(item.recv), CLng(item.qty), CLng(item.del_qty), False
                            
                            
                            
                            this_psh.Cells(x, 10) = p.pusname
                            this_psh.Cells(x, 11) = p.duns
                            this_psh.Cells(x, 12) = p.supp_nm
                            this_psh.Cells(x, 13) = p.on_mgo
                            this_psh.Cells(x, 14) = p.on_wizard
                        End If
                    End If
                Next item
            End If
            
            
        ' ------------------------------------------
    End If
End Sub

Private Sub def_recv_status(ByRef r As Range, _
    mp As pus, _
    i As PnInPus, _
    lng_recv As Long, lng_qty As Long, long_del_qty As Long, _
    Optional dopasowanie_po_nazwie As Boolean)
    
    
    
    If (CStr(r) = Cow.G_RECV_STR_BOOKED) Or (CStr(r) = Cow.G_RECV_STR_INTRANSIT) Then
    Else
    
        If dopasowanie_po_nazwie Then
        
        
            If CDate(mp.sdate) > CDate(Now) Then
                mp.recv = PU_DATE_IN_FUTURE
                r = Cow.G_RECV_STR_PU_DATE_IN_FUTURE
                
                
            ElseIf (i.content_source = CONTENT_FROM_HISTORY) And (CLng(lng_recv) > 0) Then
            
                ' zatem tylko recv
                
                If CLng(lng_recv) = CLng(lng_qty) Then
                    mp.recv = INLINE_WITH_QTY
                    r = Cow.G_RECV_STR_BOOKED
                Else
                    mp.recv = NOT_INLINE_WITH_QTY
                    r = Cow.G_RECV_STR_BOOKED_NOT_INLINE
                End If
                
            ElseIf (i.content_source = CONTENT_FROM_ASN) And (CLng(long_del_qty) > 0) Then ' Or (i.content_source = CONTENT_FROM_PUS) Then
                
                ' dane tylko z po400
                If CLng(lng_qty) = CLng(long_del_qty) Then
                    mp.recv = INLINE_WITH_QTY
                    r = Cow.G_RECV_STR_INTRANSIT
                Else
                    mp.recv = NOT_INLINE_WITH_QTY
                    r = Cow.G_RECV_STR_INTRANSIT_BUT_DIFF_QTY
                End If
                
            
            
            
            
            ElseIf i.content_source = CONTENT_FROM_WIZARD Then
                MsgBox "nie powinno znalezc dopasowania z drugim pusem z wizarda!"
                End
            Else
                mp.recv = ONLY_PUS_NAME_MATCH
                r = Cow.G_RECV_STR_ONLY_PUS_NAME_MATCH
                
                If i.content_source = CONTENT_FROM_HISTORY Then
                    r = r & ", check History Screen"
                ElseIf i.content_source = CONTENT_FROM_ASN Then
                    r = r & ", check MS9PO400 Screen"
                End If
            End If
            
            'ElseIf CLng(lng_recv) = CLng(lng_qty) And CLng(lng_recv) > 0 Then
            '
            '    mp.recv = INLINE_WITH_QTY
            '    r = Cow.G_RECV_STR_BOOKED
            '
            'ElseIf CLng(lng_qty) = CLng(long_del_qty) And CLng(lng_qty) > 0 Then
           '
           '     mp.recv = IN_TRANSIT
           '     r = Cow.G_RECV_STR_INTRANSIT
           '
           '
           ' ElseIf (CLng(lng_recv) <> CLng(lng_qty)) And (CLng(lng_qty) <> CLng(long_del_qty)) And (CLng(lng_recv) > 0 And CLng(lng_qty) > 0) Then
           '
           '     mp.recv = NOT_INLINE_WITH_QTY
           '     r = Cow.G_RECV_STR_BOOKED_NOT_INLINE
           '
           ' ElseIf ((CLng(lng_qty) <> CLng(long_del_qty)) And (CLng(lng_qty) > 0) And (CLng(lng_recv) = 0)) Then
           '
           '
           '     mp.recv = NOT_INLINE_WITH_QTY
           '     r = Cow.G_RECV_STR_INTRANSIT_BUT_DIFF_QTY
           '
           ' ElseIf CLng(lng_recv) < 0 And CLng(long_del_qty) = 0 Then
           '
           '     mp.recv = NEGATIVE_RECV
           '     r = Cow.G_RECV_STR_NEGATIVE_RECV
           '
           '
            'ElseIf dopasowanie_po_nazwie Then
            '
            ''    mp.recv = ONLY_PUS_NAME_MATCH
            '    r = Cow.G_RECV_STR_ONLY_PUS_NAME_MATCH
           '
            '
            'Else
           '
           '     mp.recv = RECV_TBD
           '     r = Cow.G_RECV_STR_TBD
           ' End If
        Else
        
            ' nie ma dopasowanie po nazwie pus
            ' zatem ja bym sie narazie skupil tylko na asn missing i juz
            If (CStr(mp.pusname) Like "*PTA*") Or (Not (CStr(mp.pusname) Like "*FMA*")) Then
                mp.recv = PTA
                r = Cow.G_RECV_STR_PTA
            Else
                mp.recv = ASN_MISSING
                r = Cow.G_RECV_STR_ASN_MISSING
            End If
        End If
    End If
End Sub



Private Function takie_same_4_od_konca(s1, s2) As Boolean
    If Right(s1, 4) = Right(s2, 4) Then
        takie_same_4_od_konca = True
    Else
        takie_same_4_od_konca = False
    End If
End Function

' ===================================================================
' ===================================================================
' ===================================================================

Public Sub take_puses_from_wizard(psh As Worksheet)


    

    Dim r As Range
    Set r = psh.Cells(2, Cow.O_INDX)
    Do
    
        If r <> "" Then
        
            If (CStr(G_FUP_CODE) = "") Or (CStr(psh.Cells(r.Row, Cow.O_FUP_code)) = CStr(G_FUP_CODE)) Then
                
                ki = CStr(psh.Cells(r.Row, Cow.O_PUS_Number))
                If puses.Exists(ki) Then
                    ' tylko dodajemy kolejny pn
                    Set item = Nothing
                    Set item = New PnInPus
                    
                    item.pn = psh.Cells(r.Row, Cow.O_PN)
                    item.qty = psh.Cells(r.Row, Cow.O_Pick_up_Qty)
                    item.del_qty = 0
                    item.recv = 0
                    
                    Set p = puses(ki)
                    p.pns.Add item
                Else
                    ' nowy pus w ogole
                    
                    
                    Set item = Nothing
                    Set item = New PnInPus
                    
                    item.pn = psh.Cells(r.Row, Cow.O_PN)
                    item.qty = psh.Cells(r.Row, Cow.O_Pick_up_Qty)
                    item.del_qty = 0
                    item.recv = 0
                    
                    item.content_source = CONTENT_FROM_WIZARD
                    
                    
                    Set p = Nothing
                    Set p = New pus
                    
                    With p
                    
                        .plt = znajdz_plt_w_arkuszu_details(psh.Parent)
                        .duns = znajdz_duns_w_arkuszu_master(psh.Parent, CStr(item.pn))
                        .supp_nm = znajdz_supp_nm_w_arkuszu_master(psh.Parent, CStr(item.pn))
                    
                        .eda = CDate(psh.Cells(r.Row, Cow.O_Delivery_Date))
                        .pusname = CStr(ki)
                        .recv = RECV_TBD
                        
                        .on_mgo = False
                        .on_wizard = True
                        
                        .fup_code = CStr(psh.Cells(r.Row, Cow.O_FUP_code))
                        
                        .sdate = CDate(psh.Cells(r.Row, Cow.O_Pick_up_date))
                        .pns.Add item
                    End With
                    
                    puses.Add CStr(ki), p
                End If
            End If
        End If
        
        Set r = special_offset(r)
    Loop Until r.Row = Cow.LAST_ROW_IN_SH
End Sub

Private Function znajdz_plt_w_arkuszu_details(w As Workbook) As String
    znajdz_plt_w_arkuszu_details = ""
    
    znajdz_plt_w_arkuszu_details = CStr(w.Sheets(Cow.DETAILS_SH_NM).Cells(Cow.plt, 2))
End Function

Private Function znajdz_duns_w_arkuszu_master(w As Workbook, strpn As String) As String
    znajdz_duns_w_arkuszu_master = ""
    
    Dim rng As Range
    Set rng = w.Sheets(Cow.MASTER_SH_NM).Cells(2, Cow.pn)
    
    If rng.Offset(1, 0) <> "" Then
        Set rng = w.Sheets(Cow.MASTER_SH_NM).Range(rng, rng.End(xlDown))
    End If
    
    
    Dim ir As Range, found As Range
    Set found = Nothing
    For Each ir In rng
        If CStr(Trim(ir)) = CStr(strpn) Then
            Set found = ir
            Exit For
        End If
    Next ir
    
    If found Is Nothing Then
        znajdz_duns_w_arkuszu_master = ""
        Exit Function
    End If
    
    znajdz_duns_w_arkuszu_master = w.Sheets(Cow.MASTER_SH_NM).Cells(found.Row, duns)
    
End Function

Private Function znajdz_supp_nm_w_arkuszu_master(w As Workbook, strpn As String) As String
    znajdz_supp_nm_w_arkuszu_master = ""
    
    Dim rng As Range
    Set rng = w.Sheets(Cow.MASTER_SH_NM).Cells(2, Cow.pn)
    
    If rng.Offset(1, 0) <> "" Then
        Set rng = w.Sheets(Cow.MASTER_SH_NM).Range(rng, rng.End(xlDown))
    End If
    
    
    Dim ir As Range, found As Range
    Set found = Nothing
    For Each ir In rng
        If CStr(Trim(ir)) = CStr(strpn) Then
            Set found = ir
            Exit For
        End If
    Next ir
    
    If found Is Nothing Then
        znajdz_supp_nm_w_arkuszu_master = ""
        Exit Function
    End If
    
    znajdz_supp_nm_w_arkuszu_master = w.Sheets(Cow.MASTER_SH_NM).Cells(found.Row, Cow.Supplier_Name)
End Function


Public Sub put_puses_into_this_puses_sh()
    
    x = CLng(first_empty_row(this_psh.Cells(1, 1)))
    For Each strki In puses.Keys()
        ' i gitara
        ' =================================================
        
        Set p = puses(strki)
        For Each item In p.pns
            
            this_psh.Cells(x, 1) = p.plt
            this_psh.Cells(x, 2) = item.pn
            this_psh.Cells(x, 3) = p.fup_code
            this_psh.Cells(x, 4) = p.sdate
            this_psh.Cells(x, 5) = p.eda
            this_psh.Cells(x, 6) = item.qty
            this_psh.Cells(x, 7) = item.del_qty
            this_psh.Cells(x, 8) = item.recv
            ' this_psh.Cells(x, 9) = p.recv
            ' Debug.Print TypeName(item)
            def_recv_status this_psh.Cells(x, Cow.RECV_STATUS_COLUMN), p, item, CLng(item.recv), CLng(item.qty), CLng(item.del_qty)
            
            
            
            this_psh.Cells(x, 10) = p.pusname
            this_psh.Cells(x, 11) = p.duns
            this_psh.Cells(x, 12) = p.supp_nm
            this_psh.Cells(x, 13) = p.on_mgo
            this_psh.Cells(x, 14) = p.on_wizard
            x = x + 1
        Next item
        
        ' =================================================
    Next
    
End Sub

Private Function first_empty_row(r As Range) As Long
    If r.Offset(1, 0) = "" Then
        first_empty_row = 2
    Else
        first_empty_row = r.End(xlDown).Row + 1
    End If
End Function


Private Sub Class_Initialize()
    Set puses = New Dictionary
End Sub


Private Sub Class_Terminate()
    Set puses = Nothing
End Sub


Public Sub pus_match_report_generator(o As Worksheet, csh As Worksheet)
    
    
    przesun_o_cbal_i_o_pot_recv o
    
    'o.Range("I1") = "POT RECV"
    'o.Range("J1") = "CBAL"
    
    wypelnij_cbalami o, csh
    
    ' to za duzo zmienia narazie usune ta logike
    ' tym bardziej ze jest ona bardzo nie spojna
    ' dodaj_do_siebie_potential_recvs o
    
End Sub

Private Sub dodaj_do_siebie_potential_recvs(o As Worksheet)

    Dim fst As Range
    Dim sec As Range
    Dim rcv As Range
    Dim pot_rcv As Range
    Dim recv_status As Range
    
    
    
    Set fst = o.Range("A2")
    Set sec = fst.Offset(1, 0)
    
    Do
        If sec & sec.Offset(0, 1) = fst & fst.Offset(0, 1) Then
        
            ' H - recv
            Set rcv = o.Range("H" & sec.Row)
            Set pot_rcv = o.Range("I" & fst.Row)
            Set recv_status = o.Range("K" & fst.Row)
            
            If rcv > 0 And pot_rcv = 0 Then
                pot_rcv = rcv
                recv_status = Cow.G_RECV_STR_POT_RECV
                ' Exit Do
            End If
            
            
            
            Set sec = sec.Offset(1, 0)
        End If
        
        If sec & sec.Offset(0, 1) <> fst & fst.Offset(0, 1) Then
            Set fst = sec
            Set sec = fst.Offset(1, 0)
        End If
        
    Loop Until sec = ""
End Sub

Private Sub wypelnij_cbalami(o As Worksheet, csh As Worksheet)

    Dim r As Range, cbal_row As Range
    ' zaczynamy od pierwszej wartosci w arkuszu pus match
    Set r = o.Range("A2")
    
    
    Do
        ' body
        ' ==========================================
        ''
        '
        ' 9 offset of pn do cbal dla arkusza pus match
        Set cbal_row = r.Offset(0, 9)
        cbal_row = CLng(znajdz_wart_cbala_dla(r, csh))
        '
        ''
        ' ==========================================
    
        Set r = r.Offset(1, 0)
    Loop Until Trim(CStr(r)) = ""
End Sub

Private Function znajdz_wart_cbala_dla(ByRef r As Range, ByRef c As Worksheet) As Long
    znajdz_wart_cbala_dla = 0
    
    pltpn = CStr(r) & CStr(r.Offset(0, 1))
    
    Dim rrr As Range
    ' arkusz cbal
    Set rrr = c.Range("A2")
    Do
        If CStr(pltpn) = CStr(Trim(rrr)) & CStr(Trim(rrr.Offset(0, 1))) Then
            
            If IsNumeric(rrr.Offset(0, 2)) Then
                znajdz_wart_cbala_dla = CLng(rrr.Offset(0, 2))
                Exit Function
            End If
        End If
        Set rrr = rrr.Offset(1, 0)
    Loop Until Trim(rrr) = ""
End Function

Private Sub przesun_o_cbal_i_o_pot_recv(ByRef o As Worksheet)
    
    If Trim(o.Range("a2")) <> "" Then
        lr = o.Range("a1").End(xlDown).Row
    Else
        lr = 2
    End If
    
    Dim rng As Range
    Set rng = o.Range("I1:O" & CStr(lr))
    rng.Copy rng.Offset(0, 2)
    
    o.Columns("I:J").Clear
    
    o.Range("I1") = "POT RECV"
    o.Range("J1") = "CBAL"
End Sub
