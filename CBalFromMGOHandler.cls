VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBalFromMGOHandler"
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




Private m As MGO

Private pnl As PN_List

Private csh As Worksheet
Private ish As Worksheet


Private Sub Class_Initialize()
    Set m = New MGO
    Set pnl = New PN_List
    
    If m.activeSession Is Nothing Then
        MsgBox "no active session on mgo!"
        End
    End If
    
    If m.getString(1, 22, 1) = "G" Then
        MsgBox "jestes niezalogowany na extra system!"
        End
    End If
    
    Set csh = ThisWorkbook.Sheets(Cow.CBALS_SH_NM)
    Set ish = ThisWorkbook.Sheets(Cow.INPUT_SH_NM)
End Sub

Private Sub Class_Terminate()
    Set m = Nothing
    Set pnl = Nothing
End Sub

Private Sub ICBalFromHandler_pobierzCbale(ich As InitConfigHandler, Optional w As Workbook)



    Dim sh As StatusHandler
    Set sh = New StatusHandler
    

    With pnl
        .pobierz_plt_raz
        .skompletuj_pny_z_arkusza_puses
        .dodaj_pny_z_arkusza_input
    End With
    
    sh.init_statusbar pnl.pn_dic.COUNT
    sh.show
    
    
    If pnl.pn_dic.COUNT > 0 Then
    
        x = 2
        For Each kii In pnl.pn_dic.Keys()
        
            With m.pMS9POP00
            
                m.sendKeys "<Clear>"
                m.sendKeys "ms9pop00 <Enter>"
                
                .DS = "6"
                .plt = pnl.plt
                .pn = kii
                m.submit
                
                csh.Cells(x, CBAL_PLT) = pnl.plt
                csh.Cells(x, CBAL_PN) = kii
                csh.Cells(x, CBAL_CBAL) = .cbal
                csh.Cells(x, Cow.CBAL_DUNS) = .duns
                csh.Cells(x, Cow.CBAL_FUP_CODE) = .F_U
                csh.Cells(x, Cow.CBAL_SUPP_NM) = .SUPPLIER
                
                x = x + 1
            End With
            
            sh.progress_increase
        Next
    End If
    
    sh.hide
    Set sh = Nothing
End Sub
