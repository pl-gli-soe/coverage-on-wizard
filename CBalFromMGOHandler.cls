VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBalFromMGOHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


    With pnl
        .pobierz_plt_raz
        .skompletuj_pny_z_arkusza_puses
        .dodaj_pny_z_arkusza_input
    End With
    
    
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
        Next
    End If
End Sub
