VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RqmsFromWizard"
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
        
        Dim sh As StatusHandler
        Set sh = New StatusHandler
        sh.init_statusbar pnl.pn_dic.COUNT
        sh.show
        
        Dim f As Range
        wiersz_outputu = 2
        For Each kii In pnl.pn_dic.Keys()
            
            Set r = pn_range(master_sheet_from_source)
            
            Set f = r.Find(CStr(kii))
            
            With this_sh
                
            
            
                If Not f Is Nothing Then
                    zrob_iteracje_pokolejnych_rqmach wiersz_outputu, f
                    wiersz_outputu = wiersz_outputu + 1
                End If
            End With
            
            
            
            
            sh.progress_increase
        Next
        
        sh.hide
        Set sh = Nothing
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
    
    
    wo = wo - 1
    
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


