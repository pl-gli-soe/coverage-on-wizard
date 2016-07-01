VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PusFromMGOHandler"
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
Implements IPUSTaker


Private w As Workbook
Private m As MGO

Private pnl As PN_List

Private psh As Worksheet
Private ish As Worksheet
Private ph As PusHandler

Private p As pus


Private ich As InitConfigHandler

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
    
    Set psh = ThisWorkbook.Sheets(Cow.PUSES_SH_NM)
    Set ish = ThisWorkbook.Sheets(Cow.INPUT_SH_NM)

End Sub

Private Sub Class_Terminate()
    Set m = Nothing
    Set pnl = Nothing
    
End Sub

Private Sub IPUSTaker_pobierzPusy(mich As InitConfigHandler, Optional s As String)


    Set ich = mich
    
    m_pobierz_pusy
End Sub

Private Sub m_pobierz_pusy()

    With pnl
        .pobierz_plt_raz
        .skompletuj_pny_z_arkusza_puses
        .dodaj_pny_z_arkusza_input
    End With
    
        
    
    With m.pZK7PPUS0
    
        Set ph = New PusHandler
        
        ' Debug.Print pn_dic.COUNT
        Dim sh As StatusHandler
        Set sh = New StatusHandler
        sh.init_statusbar pnl.pn_dic.COUNT
        sh.show
    
        For Each ki In pnl.pn_dic.Keys()
        
        
            ' wzryucam w ogole sciaganie danzch z ppus0
            If False Then
        
                If CStr(UCase(m.actualScreen)) <> "ZK7PPUS0" Then
            
                    m.sendKeys "<Clear>"
                    m.sendKeys "zk7ppus0 <Enter>"
                
                End If
            
                
                .plt = pnl.plt
                .pn = "        "
                .pn = ki
                m.submit
                
                m.sendKeys "<pf12>"
                
                ' wait and then
                If CStr(UCase(m.actualScreen)) = "ZK7PPUM0" Then
                    
                    With m.pZK7PPUM0
                    
                        ph.take_puses_from_mgo m, m.pZK7PPUM0, CStr(ki)
                    End With
                End If
            Else
            
            
            
                m.sendKeys "<Clear>"
                m.sendKeys "ms9po400 <Enter>"
                
                With m.pMS9PO400
                    .plt = pnl.plt
                    .pn = "        "
                    .pn = ki
                    m.submit
                    
                    ph.take_asns_from_mgo m, m.pMS9PO400
                End With
            End If
            
            sh.progress_increase
        Next
        
        sh.hide
        Set sh = Nothing
        
        ' tutaj jest lepiej po petli - uzupelniamy dictionary
        ' wszystkim bez polowicznych iteracji
        ' i potem ino mergujemy dane wspolnie
        ph.refresh_puses_sheet
        
        If ich.addRecv Then
            add_recv
        End If
        
        If ich.pusFromWhere = E_PUS_MIX Then
            
            ph.merge_puses_into_this_puses_sh
        ElseIf ich.pusFromWhere = E_PUS_MGO Then
        
            ph.put_puses_into_this_puses_sh
        End If
        
        
        Set ph = Nothing
        
    End With

End Sub


Private Sub add_recv()


    m.sendKeys "<Clear>"
    m.sendKeys "ms9ph100 <Enter>"
    
    Dim i As PnInPus
    
    If Trim(UCase(m.actualScreen())) = "MS9PH100" Then
    
    
        If pnl.pn_dic.COUNT > 0 Then
        
        
        
            Dim sh As StatusHandler
            Set sh = New StatusHandler
            sh.init_statusbar pnl.pn_dic.COUNT
            sh.show
        
        
            On Error Resume Next
            For Each ki In pnl.pn_dic.Keys()
            
                With m.pMS9PH100
                    .plt = pnl.plt
                    ' .pn = "        "
                    .pn = ki
                    .DATE_FROM = m.convertToMgoShortDate(CDate(Date - Cow.ILE_DNI))
                    m.submit
                    
                    ' I4017: PART NOT FOUND
                    If m.getCode("ms9ph100") <> "I4017" Then
                        y = 0
                        Do
                            For x = y To 9
                                ' Debug.Print ph.get_puses().COUNT
                                If .DTYP(Int(x)) = "RECV" Then
                                    If ph.get_puses().Exists(Trim(CStr(.SID(Int(x))))) Then
                                        
                                        ' match section
                                        ' =============================================================
                                        
                                        Set p = ph.get_puses()(.SID(Int(x)))
                                        
                                        Dim czy_znalazlem_pn As Boolean
                                        czy_znalazlem_pn = False
                                        For Each i In p.pns
                                        
                                            ' mogla byc tak ze przedtem mialem match na poziomie Cstr
                                            ' i na ekranie historii bylo to dodatkowe zero po lewej stronie co
                                            ' moglo zaciemniac obraz
                                            If CLng(i.pn) = CLng(ki) Then
                                                czy_znalazlem_pn = True
                                                If IsNumeric(.qty(Int(x))) Then
                                                
                                                    ' no bardzo wazny element pozbywajacy sie przecinka dla wartosci
                                                    ' wiekszych niz 1k
                                                    i.recv = CLng(Replace(CStr(.qty(Int(x))), ",", ""))
                                                    
                                                    i.content_source = CONTENT_FROM_HISTORY
                                                    
                                                    
                                                    If CLng(.qty(Int(x))) = 0 Then
                                                        p.recv = ON_ZERO
                                                    ElseIf CLng(i.recv) = CLng(i.qty) Then
                                                        p.recv = INLINE_WITH_QTY
                                                    Else
                                                        p.recv = NOT_INLINE_WITH_QTY
                                                    End If
                                                Else
                                                    i.recv = 0
                                                    p.recv = ON_ZERO
                                                End If
                                            End If
                                        Next i
                                        
                                        If czy_znalazlem_pn = False Then
                                            
                                            Set i = New PnInPus
                                            i.del_qty = 0
                                            i.pn = CStr(ki)
                                            i.qty = 0
                                            
                                            If IsNumeric(.qty(Int(x))) Then
                                            
                                                i.recv = CLng(Replace(CStr(.qty(Int(x))), ",", ""))
                                                
                                                i.content_source = CONTENT_FROM_HISTORY
                                                
                                                If CLng(.qty(Int(x))) = 0 Then
                                                    p.recv = ON_ZERO
                                                ElseIf CLng(i.recv) = CLng(i.qty) Then
                                                    p.recv = INLINE_WITH_QTY
                                                Else
                                                    p.recv = NOT_INLINE_WITH_QTY
                                                End If
                                            Else
                                                i.recv = 0
                                                p.recv = ON_ZERO
                                            End If
                                            
                                            p.pns.Add i
                                        End If
                                        
                                        ' =============================================================
                                    Else
                                        
                                        Set i = Nothing
                                        Set i = New PnInPus
                                        
                                        i.pn = m.pMS9PH100.pn
                                        i.recv = CLng(Replace(CStr(.qty(Int(x))), ",", ""))
                                        
                                        i.content_source = CONTENT_FROM_HISTORY
                                        
                                        
                                        Set p = Nothing
                                        Set p = New pus
                                        
                                        With p
                                        
                                            .plt = m.pMS9PH100.plt
                                            .duns = m.pMS9PH100.duns(Int(x))
                                            .supp_nm = ""
                                        
                                            .eda = CDate(m.convertToDateFromMS9PH100ShortDate(Trim(m.pMS9PH100.RECV_DATE(Int(x)))))
                                            .pusname = Trim(CStr(m.pMS9PH100.SID(Int(x))))
                                            
                                            .recv = RECV_TBD
    
                                            
                                            .on_mgo = True
                                            .on_wizard = False
                                            
                                            .fup_code = ""
                                            
                                            ' .sdate = CDate(m.convertToDateFromZK7PPUM0Date(Trim(scr.PU_DATE(Int(x)))))
                                            .pns.Add i
                                        End With
                                        
                                        ph.get_puses().Add CStr(p.pusname), p
                                        
                                    End If
                                End If
                            Next x
                            
                            If m.getCode("ms9ph100") = "R6086" Then
                                Exit Do
                            End If
                            
                            m.sendKeys "<pf8>"
                            y = 1
                            
                        Loop While True
                    End If
                    
                End With
                
                sh.progress_increase
            Next ki
            
            sh.hide
            Set sh = Nothing
        End If
    End If
    
End Sub

