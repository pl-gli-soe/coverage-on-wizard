VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBalFromWizardHandler"
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


Implements ICBalFromHandler



Private pnl As PN_List

Private csh As Worksheet
Private ish As Worksheet

Private master_sheet_from_source As Worksheet
Private m As Worksheet


Private Sub Class_Initialize()
    Set pnl = New PN_List
    
    
    Set csh = ThisWorkbook.Sheets(Cow.CBALS_SH_NM)
    Set ish = ThisWorkbook.Sheets(Cow.INPUT_SH_NM)
    
    
    Set master_sheet_from_source = G_SOURCE_WIZARD.Sheets(Cow.MASTER_SH_NM)
    Set m = master_sheet_from_source
End Sub

Private Sub Class_Terminate()
    Set pnl = Nothing
End Sub

Private Sub ICBalFromHandler_pobierzCbale(ich As InitConfigHandler, Optional w As Workbook)



    
     
    

    With pnl
        .pobierz_plt_raz
        .skompletuj_pny_z_arkusza_puses
        .dodaj_pny_z_arkusza_input
    End With
    
    Dim sh As StatusHandler
    Set sh = New StatusHandler
    sh.init_statusbar pnl.pn_dic.COUNT
    sh.show
    
    Dim r As Range
    Dim moj_cbal As Long
    
    
    If pnl.pn_dic.COUNT > 0 Then
    
        x = 2
        For Each kii In pnl.pn_dic.Keys()
        
            Set r = m.Cells(2, Cow.pn)
            Do
            
                If CStr(kii) = CStr(Trim(r)) Then
        
                    moj_cbal = CLng(znajdz_kolumne_cbal_w_wizardzie(r))
                    
                    csh.Cells(x, CBAL_PLT) = pnl.plt
                    csh.Cells(x, CBAL_PN) = kii
                    csh.Cells(x, CBAL_CBAL) = moj_cbal
                    csh.Cells(x, Cow.CBAL_DEL_CONF) = m.Cells(r.Row, Cow.Delivery_confirmation)
                    csh.Cells(x, Cow.CBAL_DUNS) = m.Cells(r.Row, Cow.duns)
                    csh.Cells(x, Cow.CBAL_RESP) = m.Cells(r.Row, Cow.Responsibility)
                    csh.Cells(x, Cow.CBAL_FUP_CODE) = m.Cells(r.Row, Cow.fup_code)
                    csh.Cells(x, Cow.CBAL_SUPP_NM) = m.Cells(r.Row, Cow.Supplier_Name)
                    csh.Cells(x, Cow.CBAL_TOTAL_MRD_QTY) = m.Cells(r.Row, Cow.Total_QTY)
                    csh.Cells(x, Cow.CBAL_TOTAL_PUS_QTY) = m.Cells(r.Row, Cow.total_pus_qty)
                    
                    x = x + 1
                End If
                
                Set r = r.Offset(1, 0)
            Loop Until Trim(r) = ""
            
            sh.progress_increase
        Next
    End If
    
    sh.hide
    Set sh = Nothing
End Sub

Private Function znajdz_kolumne_cbal_w_wizardzie(ir As Range) As Long
    znajdz_kolumne_cbal_w_wizardzie = 0
    
    Dim tmp As Range
    Set tmp = ir.Parent.Cells(1, 1)
    Do
    
        If tmp = Cow.G_STR_CBAL Then
            If IsNumeric(ir.Parent.Cells(ir.Row, tmp.Column)) Then
                znajdz_kolumne_cbal_w_wizardzie = CLng(ir.Parent.Cells(ir.Row, tmp.Column))
            Else
                
            End If
            
            Exit Function
        End If
        Set tmp = tmp.Offset(0, 1)
    Loop Until Trim(tmp) = ""
    
End Function
