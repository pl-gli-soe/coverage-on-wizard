VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoordList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements ICoverage

Private puses As Worksheet
Private cbals As Worksheet

Private coordsh As Worksheet
Private r As Range

Private slownik As Dictionary
Private ci As CoverageIteration


Private puses_from_where As E_TYPE_OF_PUSES_FOR_COVERAGE



Private Sub Class_Initialize()
    Set puses = ThisWorkbook.Sheets(Cow.PUSES_SH_NM)
    Set cbals = ThisWorkbook.Sheets(Cow.CBALS_SH_NM)
    
    
    Set coordsh = ThisWorkbook.Sheets.Add
    
    
    Set slownik = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set slownik = Nothing
End Sub

Private Sub ICoverage_makeCoverage(ich As InitConfigHandler)


    puses_from_where = ich.pusesForCoverage


    Set r = coordsh.Range("a1")
    r = "COORD LIST BASED ON: , " & Cow.PUSES_SH_NM & " & " & Cow.CBALS_SH_NM & ". Time: " & CStr(Now)
    
    Set r = r.Offset(1, 0)

    pobierz_cbale
    pobierz_pusy
    
    przygotuj_coord_list_na_postawie_pobranych_danych coordsh
    simpleAutoFit
    koloruj
End Sub

Private Sub simpleAutoFit()

    With coordsh
        
        .Columns("B:ZZ").AutoFit
    End With
End Sub

Private Sub przygotuj_coord_list_na_postawie_pobranych_danych(s As Worksheet)


    ' r jest ustawiony juz na a2
    ' ==============================================================
    Set r = s.Range("A2")
    r = "PLT"
    r.Offset(0, Cow.COORDLIST_PN - 1) = "PN"
    r.Offset(0, Cow.COORDLIST_DUNS - 1) = "DUNS"
    r.Offset(0, Cow.COORDLIST_SUPPNM - 1) = "SUPP NM"
    r.Offset(0, Cow.COORDLIST_RESP - 1) = "RESP"
    r.Offset(0, Cow.COORDLIST_FUPCODE - 1) = "FUP CODE"
    r.Offset(0, Cow.COORDLIST_TOTAL - 1) = "TOTAL MRD QTY"
    
    r.Offset(0, Cow.COORDLIST_CBAL - 1) = "CBAL"
    
    r.Offset(0, Cow.COORDLIST_DELCONF - 1) = "DEL CONF"
    r.Offset(0, Cow.COORDLIST_TOTALPUS - 1) = "TOTAL PUS QTY"
    
    
    i = 1 ' ile iteracji kolumn
    
    ' A3 lecimy z danymi
    Set r = r.Offset(1, 0)
    
    For Each k In slownik.Keys()
    
        Set ci = slownik(k)
        r = ci.plt
        r.Offset(0, Cow.COORDLIST_PN - 1) = ci.pn
        r.Offset(0, Cow.COORDLIST_DUNS - 1) = ci.duns
        r.Offset(0, Cow.COORDLIST_SUPPNM - 1) = ci.supp_nm
        r.Offset(0, Cow.COORDLIST_RESP - 1) = ci.resp
        r.Offset(0, Cow.COORDLIST_FUPCODE - 1) = ci.fup_code
        r.Offset(0, Cow.COORDLIST_TOTAL - 1) = ci.total_mrd_qty
        
        r.Offset(0, Cow.COORDLIST_CBAL - 1) = ci.cbal
        
        r.Offset(0, Cow.COORDLIST_DELCONF - 1) = ci.del_conf
        r.Offset(0, Cow.COORDLIST_TOTALPUS - 1) = ci.get_total_pus_qty
        
        
        
        lecimy_z_iteracja_pusow r.Offset(0, Cow.COORDLIST_FIRST_PUS_PARAM - 1), ci, i
        
        
        Set r = r.Offset(1, 0)
        
    Next
    
    
    zbuduj_labelki_pod_max_pusow_w_iteracji i
    ' ==============================================================

End Sub

Private Sub zbuduj_labelki_pod_max_pusow_w_iteracji(i As Variant)
    
    For x = 1 To i
        coordsh.Range("A2").End(xlToRight).Offset(0, 1) = "DEL DATE"
        coordsh.Range("A2").End(xlToRight).Offset(0, 1) = "PUS DATE"
        coordsh.Range("A2").End(xlToRight).Offset(0, 1) = "QTY"
        coordsh.Range("A2").End(xlToRight).Offset(0, 1) = "PUS #"
    Next x
End Sub

Private Sub lecimy_z_iteracja_pusow(rr As Range, c As CoverageIteration, ByRef ile_iteracji_kolumn As Variant)

    Dim i As CoordItem
    
    If c.get_puses.COUNT > ile_iteracji_kolumn Then
        ile_iteracji_kolumn = c.get_puses.COUNT
    End If

    For Each k In c.get_puses().Keys()
        
        ' BODY SECTION
        ' ----------------------------------------------------------
        Set i = c.get_puses()(k)
        rr = i.pus_date
        rr.Offset(0, 1) = i.del_date
        rr.Offset(0, 2) = i.qty
        rr.Offset(0, 3) = i.pus_name
        
        Set rr = rr.Offset(0, 4)
        
        ' ----------------------------------------------------------
    Next
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
                ci.dodaj_lub_zmien_pus_dla_coord_list puses.Cells(r.Row, Cow.PUSES_PUS_NAME), r, puses_from_where
                
                ' -----------------------------------------------------------------------------------------
            End If
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    ' -----------------------------------------------------------------------------------------------------
End Sub

Private Sub koloruj()

End Sub
