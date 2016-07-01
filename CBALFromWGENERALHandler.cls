VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBALFromWGENERALHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2016 Mateusz Forrest Milewski
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

Private wgen As Workbook
Private wgen_worksheet As Worksheet


Private Sub Class_Initialize()
    Set pnl = New PN_List
    
    
    Set csh = ThisWorkbook.Sheets(Cow.CBALS_SH_NM)
    Set ish = ThisWorkbook.Sheets(Cow.INPUT_SH_NM)
    
    
    
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
    
    Set wgen = w
    
    
    If Not wgen Is Nothing Then
    
    Set wgen_worksheet = wgen.ActiveSheet
    
        Dim sh As StatusHandler
        Set sh = New StatusHandler
        sh.init_statusbar pnl.pn_dic.COUNT
        sh.show
        
        
        Dim r As Range
        Dim moj_cbal As Long
        
        
        If pnl.pn_dic.COUNT > 0 Then
        
            x = 2
            For Each kii In pnl.pn_dic.Keys()
            
                ' I col: FUP, II col: PLT, III col: PN
                Set r = wgen_worksheet.Cells(2, 2)
                Do
                
                    If (CLng(kii) = CLng(Trim(r.Offset(0, 1)))) And (pnl.plt = Trim(r)) Then
            
                        moj_cbal = CLng(znajdz_kolumne_cbal_wgen(r))
                        '
                        csh.Cells(x, CBAL_PLT) = pnl.plt
                        csh.Cells(x, CBAL_PN) = kii
                        csh.Cells(x, CBAL_CBAL) = moj_cbal
                        
                        csh.Cells(x, Cow.CBAL_DUNS) = wgen_worksheet.Cells(r.Row, 13) ' DUNS
                        'csh.Cells(x, Cow.CBAL_RESP) = m.Cells(r.Row, Cow.Responsibility)
                        csh.Cells(x, Cow.CBAL_FUP_CODE) = wgen_worksheet.Cells(r.Row, 1) ' FUP CODE
                        csh.Cells(x, Cow.CBAL_SUPP_NM) = wgen_worksheet.Cells(r.Row, 14) ' SUPP NAME
                        'csh.Cells(x, Cow.CBAL_TOTAL_MRD_QTY) = m.Cells(r.Row, Cow.Total_QTY)
                        'csh.Cells(x, Cow.CBAL_TOTAL_PUS_QTY) = m.Cells(r.Row, Cow.total_pus_qty)
                        
                        x = x + 1
                    End If
                    
                    Set r = r.Offset(1, 0)
                Loop Until Trim(r) = ""
                
                sh.progress_increase
            Next
        End If
        
        sh.hide
        Set sh = Nothing
    End If
End Sub

Private Function znajdz_kolumne_cbal_wgen(r As Range) As Long
    znajdz_kolumne_cbal_wgen = 0
    
    ' cstr = "Beginning On Hand Qty"
    
    Dim ir As Range
    Set ir = r.Parent.Cells(r.Row, 5) ' BEG_ON_HAND_COLUMN_FROM_GENERAL
    
    znajdz_kolumne_cbal_wgen = CLng(ir)
    
End Function
