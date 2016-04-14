VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Coverage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements ICoverage


Private rqms As Worksheet
Private puses As Worksheet
Private cbals As Worksheet

Private covsh As Worksheet
Private r As Range

Private slownik As Dictionary
Private ci As CoverageIteration


Private puses_from_where As E_TYPE_OF_PUSES_FOR_COVERAGE



Private Sub ICoverage_makeCoverage(ich As InitConfigHandler)


    puses_from_where = ich.pusesForCoverage


    Set r = covsh.Range("a1")
    r = "coverage based on: " & Cow.RQMS_SH_NM & ", " & Cow.PUSES_SH_NM & " & " & Cow.CBALS_SH_NM & ". Time: " & CStr(Now)
    
    Set r = r.Offset(1, 0)

    pobierz_cbale
    pobierz_rqmsy
    pobierz_pusy
    
    przygotuj_coverage_na_postawie_pobranych_danych
    koloruj
End Sub

Private Sub Class_Initialize()

    Set rqms = ThisWorkbook.Sheets(Cow.RQMS_SH_NM)
    Set puses = ThisWorkbook.Sheets(Cow.PUSES_SH_NM)
    Set cbals = ThisWorkbook.Sheets(Cow.CBALS_SH_NM)
    
    
    Set covsh = ThisWorkbook.Sheets.Add
    
    
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
                ci.cbal = CLng(r.Offset(0, 2))
                
                
                
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

End Sub

Private Sub koloruj()

End Sub


