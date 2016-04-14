VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoverageIteration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    
            If slownik_cov_items.Exists(cwkey) Then
                
                Set i = slownik_cov_items(cwkey)
                i.pus = i.pus + CLng(tmp)
                            
            Else
                Set i = New CoverageItem
                i.ycw = CStr(cwkey)
                i.pus = CLng(tmp)
                
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
