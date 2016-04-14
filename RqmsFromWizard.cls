VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RqmsFromWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IRqmTaker


Private master_sheet_from_source As Worksheet
Private details_sheet_from_source As Worksheet
Private cmnt_src_sheet_from_source As Worksheet
Private cs As Worksheet


Private this_sh As Worksheet

Private pnl As PN_List
Private ki As String

Private r As Range

Private kolekcja_itemow_range_dla_jednego_pn_zawierajacych_rqmsy As Collection


Private Sub Class_Initialize()
    Set pnl = New PN_List
End Sub

Private Sub Class_Terminate()
    Set pnl = Nothing
End Sub

Private Sub IRqmTaker_pobierzRqmsy(mich As InitConfigHandler, w As Workbook)


    Set master_sheet_from_source = w.Sheets(Cow.MASTER_SH_NM)
    Set details_sheet_from_source = w.Sheets(Cow.DETAILS_SH_NM)
    Set cmnt_src_sheet_from_source = w.Sheets(Cow.COMMENT_SOURCE_SH_NM)
    Set cs = cmnt_src_sheet_from_source
    Set this_sh = ThisWorkbook.Sheets(Cow.RQMS_SH_NM)
    
    With pnl
        .pobierz_plt_raz
        .skompletuj_pny_z_arkusza_puses
        .dodaj_pny_z_arkusza_input
    End With
    
    If pnl.pn_dic.COUNT > 0 Then
    
        ' RQMS downloading section
        ' ==================================================
        Dim f As Range
        wiersz_outputu = 2
        For Each kii In pnl.pn_dic.Keys()
            
            Set r = pn_range(master_sheet_from_source)
            
            Set f = r.Find(CStr(kii))
            
            With this_sh
                
            
            
                If Not f Is Nothing Then
                    zrob_iteracje_pokolejnych_rqmach wiersz_outputu, f
                Else
                    .Cells(wiersz_outputu, Cow.RQMS_PLT) = pnl.plt
                    .Cells(wiersz_outputu, Cow.RQMS_PN) = ki
                    .Cells(wiersz_outputu, Cow.RQMS_CW) = ""
                    .Cells(wiersz_outputu, Cow.RQMS_QTY) = 0
                End If
            End With
            
            
            wiersz_outputu = wiersz_outputu + 1
            
        Next
        
        ' ==================================================
    End If

End Sub

Private Function pn_range(sh As Worksheet) As Range
    
    Set pn_range = Nothing
    
    Set pn_range = sh.Cells(2, Cow.pn)
    
    If Trim(pn_range.Offset(1, 0)) <> "" Then
        Set pn_range = sh.Range(pn_range, pn_range.End(xlDown))
    End If
    
End Function

Private Sub zrob_iteracje_pokolejnych_rqmach(ByRef wo As Variant, ByRef f As Range)
    
    ' sekcja kolejnych iteracji dla rqmsow
    ' ===================================================================================
    
    Dim ir As Range
    Set ir = master_sheet_from_source.Cells(1, Cow.BLANK4 + 1)
    
    Do
        If Not ir.Comment Is Nothing Then
            ' jest comment to znaczy ze potencjalnie jest to kolumna wraz z wartosciami - teraz trzeba sprawdzic arkusz comment source
            ' ---------------------------------------------------------------------------------------------------------------------------
            
            ycw = ""
            v = 0
            
            ycw = znajdz_cw_dla_tego_ira(ir, cs)
            If ycw <> "" Then
                v = master_sheet_from_source.Cells(f.Row, ir.Column)
                
                this_sh.Cells(wo, Cow.RQMS_PLT) = pnl.plt
                this_sh.Cells(wo, Cow.RQMS_PN) = master_sheet_from_source.Cells(f.Row, Cow.pn)
                this_sh.Cells(wo, Cow.RQMS_CW) = ycw
                this_sh.Cells(wo, Cow.RQMS_QTY) = v
                
                
                wo = wo + 1
            End If
            ' ---------------------------------------------------------------------------------------------------------------------------
        End If
        
        Set ir = ir.Offset(0, 1)
    Loop Until Trim(ir) = ""
    
    ' ===================================================================================
End Sub

Private Function znajdz_cw_dla_tego_ira(iir As Range, ics As Worksheet) As String
    znajdz_cw_dla_tego_ira = ""
    
    Dim tmp As Range
    Set tmp = ics.Cells(1, 1)
    Do
    
        If Trim(tmp) = Trim(iir) Then
            
            Do
                If tmp Like "Y*CW*" Then
                    znajdz_cw_dla_tego_ira = tmp
                    Exit Function
                End If
            
                Set tmp = tmp.Offset(1, 0)
            Loop Until Trim(tmp) = ""
        End If
        
        Set tmp = tmp.Offset(0, 1)
    Loop Until Trim(tmp) = ""
End Function


