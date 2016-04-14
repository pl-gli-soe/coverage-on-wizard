VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WybierzPlikForm 
   Caption         =   "Wybierz Plik: "
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3045
   OleObjectBlob   =   "WybierzPlikForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WybierzPlikForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private e As E_PUS_CZY_RQM_CZY_CBAL

Public Sub init(m_e As E_PUS_CZY_RQM_CZY_CBAL)

    e = m_e
    
    Me.ListBox1.Clear
    With Me.ListBox1
        
        Dim w As Workbook
        For Each w In Workbooks
            .AddItem w.Name
        Next w
    End With
End Sub



Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    hide
    If Me.ListBox1.ListCount > 0 Then
    
    
        If e = FOMULARZ_WYBORU_PLIKU_DLA_PUS Then
        
            Dim w As IPUSTaker
            Set w = New WizardHandler
            w.pobierzPusy Nothing, Me.ListBox1.Value
        ElseIf e = FOMULARZ_WYBORU_PLIKU_DLA_RQM Then
            
            Dim r As IRqmTaker
            Set r = New RqmsFromWizard
            r.pobierzRqmsy Nothing, Workbooks(CStr(Me.ListBox1.Value))
            
        ElseIf e = FOMULARZ_WYBORU_PLIKU_DLA_CBAL Then
        
            Dim c As ICBalFromHandler
            Set c = New CBalFromWizardHandler
            c.pobierzCbale Nothing
        End If
        
    Else
        MsgBox "nie ma czego wybrac!"
    End If

End Sub
