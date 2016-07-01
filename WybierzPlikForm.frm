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

Private e As E_PUS_CZY_RQM_CZY_CBAL
Private ich As InitConfigHandler

Public Sub init(m_e As E_PUS_CZY_RQM_CZY_CBAL, Optional mich As InitConfigHandler)

    e = m_e
    Set ich = Nothing
    
    On Error Resume Next
    If IsMissing(mich) Then
        Set ich = Nothing
    Else
        Set ich = mich
    End If
    
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
            w.pobierzPusy ich, CStr(Me.ListBox1.Value)
        ElseIf e = FOMULARZ_WYBORU_PLIKU_DLA_RQM Then
            
            Dim r As IRqmTaker
            Set r = New RqmsFromWizard
            r.pobierzRqmsy Nothing, Workbooks(CStr(Me.ListBox1.Value))
            
        ElseIf e = FOMULARZ_WYBORU_PLIKU_DLA_CBAL Then
        
            Set Cow.G_SOURCE_WIZARD = Workbooks(CStr(Me.ListBox1.Value))
            Dim c As ICBalFromHandler
            Set c = New CBalFromWizardHandler
            c.pobierzCbale Nothing, Workbooks(CStr(Me.ListBox1.Value))
        End If
        
    Else
        MsgBox "nie ma czego wybrac!"
    End If

End Sub
