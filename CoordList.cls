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
    r = "coverage based on: , " & Cow.PUSES_SH_NM & " & " & Cow.CBALS_SH_NM & ". Time: " & CStr(Now)
    
    Set r = r.Offset(1, 0)

    pobierz_cbale
    pobierz_pusy
    
    przygotuj_coord_list_na_postawie_pobranych_danych coordsh
    koloruj
End Sub

Private Sub przygotuj_coord_list_na_postawie_pobranych_danych(s As Worksheet)


    ' r jest ustawiony juz na a2
    ' ==============================================================
    Set r = s.Range("A2")
    r = "PLT"
    r.Offset(0, 1) = "PN"
    r.Offset(0, 2) = "CBAL"
    
    ' A3 lecimy z danymi
    Set r = r.Offset(1, 0)
    
    For Each k In slownik.Keys()
    
        Set ci = slownik(k)
        r = ci.plt
        r.Offset(0, 1) = ci.pn
        r.Offset(0, 2) = ci.cbal
        
        lecimy_z_iteracja_pusow r.Offset(0, 3), ci
        
        Set r = r.Offset(1, 0)
        
    Next
    
    
    ' ==============================================================

End Sub

Private Sub lecimy_z_iteracja_pusow(rr As Range, c As CoverageIteration)

    Dim i As CoordItem

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
                ci.cbal = CLng(r.Offset(0, 2))
                
                
                
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
